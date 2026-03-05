/**
 * SuperDoc Document API Agent
 *
 * Connects to the collaboration server using the Document API SDK.
 * Demonstrates how an AI agent can join a collaboration session and use LLM tools.
 *
 * Usage: npx tsx agent.ts [documentId]
 */

import 'dotenv/config';
import { createSuperDocClient, chooseTools, dispatchSuperDocTool } from '@superdoc-dev/sdk';
import type { SuperDocClient } from '@superdoc-dev/sdk';
import OpenAI from 'openai';
import * as Y from 'yjs';
import { WebsocketProvider } from 'y-websocket';

const COLLAB_URL = 'ws://localhost:3050/collaboration';
const DEFAULT_DOC_ID = 'superdoc-demo';
const DEBOUNCE_MS = 1000;

// Flag to prevent infinite loop - ignore updates while agent is writing
let isAgentWriting = false;

// Buffer to track last seen content - only prompt if there's new human content
let lastSeenContent = '';

// Tools with invalid "type": "json" in schemas - OpenAI rejects these
const EXCLUDED_TOOLS = [
  'apply_mutations',
  'preview_mutations',
  'doc_mutations_apply',
  'doc_mutations_preview',
  'doc_lists_setLevelRestart',
  'doc_lists_setValue',
  'doc_sections_setPageBorders',
  'set_list_level_restart',
  'set_list_value',
  'set_section_page_borders',
];

/**
 * Create and connect to the SuperDoc host process.
 */
async function connectToSuperdoc(): Promise<SuperDocClient> {
  const client = createSuperDocClient();
  await client.connect();
  console.log('[Agent] Connected to SuperDoc host');
  return client;
}

/**
 * Join a collaboration room via SDK.
 */
async function joinCollaboration(client: SuperDocClient, documentId: string): Promise<void> {
  console.log(`[Agent] Joining collaboration room: ${documentId}`);
  await client.doc.open({
    collaboration: {
      // Needs documentation: providerType must be specified when using collaboration object.
      providerType: 'y-websocket',
      url: COLLAB_URL,
      documentId: documentId,
    },
  });
  console.log('[Agent] Successfully joined collaboration!');
}

/**
 * Connect to Yjs WebSocket for real-time document updates.
 * Returns cleanup function.
 */
function connectToYjsUpdates(
  documentId: string,
  onUpdate: () => void,
): { ydoc: Y.Doc; provider: WebsocketProvider; cleanup: () => void } {
  const ydoc = new Y.Doc();
  const provider = new WebsocketProvider(COLLAB_URL, documentId, ydoc);

  provider.on('sync', (synced: boolean) => {
    if (synced) {
      console.log('[Agent] Yjs synced with server');
    }
  });

  // Listen for remote updates (not local changes)
  ydoc.on('update', (_update: Uint8Array, origin: unknown) => {
    // Only trigger on remote changes (origin is the provider for remote updates)
    if (origin === provider) {
      onUpdate();
    }
  });

  const cleanup = () => {
    provider.disconnect();
    ydoc.destroy();
  };

  return { ydoc, provider, cleanup };
}

/**
 * Create a debounced handler with abort support.
 */
function createDebouncedHandler(
  handler: (signal: AbortSignal) => Promise<void>,
  delayMs: number,
): { trigger: () => void; cancel: () => void } {
  let timeoutId: NodeJS.Timeout | null = null;
  let abortController: AbortController | null = null;

  const cancel = () => {
    if (timeoutId) {
      clearTimeout(timeoutId);
      timeoutId = null;
    }
    if (abortController) {
      abortController.abort();
      abortController = null;
    }
  };

  const trigger = () => {
    // Cancel any pending operation
    cancel();

    // Start new debounce
    timeoutId = setTimeout(async () => {
      abortController = new AbortController();
      try {
        await handler(abortController.signal);
      } catch (error) {
        if ((error as Error).name !== 'AbortError') {
          console.error('[Agent] Handler error:', error);
        }
      }
      abortController = null;
    }, delayMs);
  };

  return { trigger, cancel };
}

/**
 * Read document and continue writing with LLM.
 */
async function readAndContinueWriting(
  client: SuperDocClient,
  openai: OpenAI,
  tools: OpenAI.ChatCompletionTool[],
  signal: AbortSignal,
): Promise<void> {
  if (signal.aborted) return;

  console.log('[Agent] User stopped typing, reading document...');

  try {
    const documentText = await client.doc.getText({});
    if (signal.aborted) return;

    const currentContent = typeof documentText === 'string' ? documentText : String(documentText);

    console.log('[Agent] Current document:');
    console.log('─'.repeat(50));
    console.log(currentContent);
    console.log('─'.repeat(50));

    // Skip if document is empty or too short
    if (currentContent.trim().length < 10) {
      console.log('[Agent] Document too short, waiting for more content...');
      return;
    }

    // Skip if no new content since last check
    if (currentContent === lastSeenContent) {
      console.log('[Agent] No new content, skipping...');
      return;
    }

    // Check if there's actually new human content (not just whitespace changes)
    const newContent = currentContent.slice(lastSeenContent.length).trim();
    if (newContent.length === 0 && currentContent.startsWith(lastSeenContent)) {
      console.log('[Agent] Only whitespace changes, skipping...');
      return;
    }

    console.log(`[Agent] New content detected: "${newContent.slice(0, 50)}${newContent.length > 50 ? '...' : ''}"`);

    const systemPrompt = `You are a creative writing assistant. Your job is to continue the user's writing in a natural, engaging way.

IMPORTANT: A document session is already open. Do NOT pass "doc" or "sessionId" parameters.

To append content, use insert_content with markdown formatting:
{ "value": "your content here", "type": "markdown" }

You can use markdown syntax in your value:
- # Heading 1, ## Heading 2, ### Heading 3
- **bold**, *italic*
- Regular paragraph text

Rules:
- Continue the story/text naturally from where it left off
- Match the tone and style of the existing writing
- Add 1-2 sentences OR a new section heading when appropriate
- Occasionally add a heading (# or ##) to start a new chapter or section
- Do NOT repeat what's already written
- Do NOT add meta-commentary, just continue the narrative`;

    const messages: OpenAI.ChatCompletionMessageParam[] = [
      { role: 'system', content: systemPrompt },
      { role: 'user', content: `Here is the current document. Continue writing from where it ends:\n\n${documentText}` },
    ];

    if (signal.aborted) return;

    const response = await openai.chat.completions.create({
      model: 'gpt-4o',
      messages,
      tools,
      tool_choice: 'required',
    });

    if (signal.aborted) return;

    const message = response.choices[0].message;

    if (message.tool_calls?.length) {
      for (const call of message.tool_calls) {
        if (signal.aborted) return;

        const args = JSON.parse(call.function.arguments);
        console.log(`[Agent] Executing tool: ${call.function.name}`, JSON.stringify(args));
        try {
          // Set flag to ignore our own updates
          isAgentWriting = true;
          await dispatchSuperDocTool(client, call.function.name, args);
          console.log('[Agent] Successfully continued the writing!');

          // Update buffer with new content (including what we just wrote)
          const updatedText = await client.doc.getText({});
          lastSeenContent = typeof updatedText === 'string' ? updatedText : String(updatedText);
          console.log('[Agent] Buffer updated with new content');
        } catch (error) {
          const errorMsg = error instanceof Error ? error.message : String(error);
          console.error(`[Agent] Tool error: ${errorMsg}`);
        } finally {
          // Small delay to let Yjs sync propagate, then re-enable listening
          setTimeout(() => {
            isAgentWriting = false;
          }, 500);
        }
      }
    } else {
      console.log('[Agent] LLM did not use a tool:', message.content);
    }
  } catch (error) {
    if (!signal.aborted) {
      console.error('[Agent] Error:', error);
    }
  }
}

/**
 * Initialize LLM tools for document editing.
 */
async function initLLMTools(): Promise<{ tools: OpenAI.ChatCompletionTool[]; openai: OpenAI }> {
  const { tools, selected } = await chooseTools({
    provider: 'openai',
    taskContext: { phase: 'mutate' },
    policy: {
      allowMutatingTools: true,
      forceExclude: EXCLUDED_TOOLS,
    },
  });

  console.log('[Agent] Available tools:', selected.map((t) => t.toolName).join(', '));
  console.log(`[Agent] Loaded ${tools.length} tools for LLM`);

  const openai = new OpenAI();
  return { tools: tools as OpenAI.ChatCompletionTool[], openai };
}

/**
 * Run the agentic loop - send messages to LLM and execute tool calls.
 */
async function runAgentLoop(
  client: SuperDocClient,
  openai: OpenAI,
  tools: OpenAI.ChatCompletionTool[],
  userPrompt: string,
): Promise<string | null> {
  const systemPrompt = `You are a document editing assistant. You MUST use the provided tools to make changes to the document.

IMPORTANT: A document session is already open. Do NOT pass "doc" or "sessionId" parameters.

To insert text at the end of the document, use insert_content with just the value parameter:
{ "value": "your text here" }

That's it - no target needed for appending to the end.`;

  const messages: OpenAI.ChatCompletionMessageParam[] = [
    { role: 'system', content: systemPrompt },
    { role: 'user', content: userPrompt },
  ];

  console.log('[Agent] Sending prompt to LLM...');

  let isFirstCall = true;
  while (true) {
    const response = await openai.chat.completions.create({
      model: 'gpt-4o',
      messages,
      tools,
      tool_choice: isFirstCall ? 'required' : 'auto',
    });
    isFirstCall = false;

    const message = response.choices[0].message;
    messages.push(message);

    if (!message.tool_calls?.length) {
      console.log('[Agent] LLM response:', message.content);
      return message.content;
    }

    for (const call of message.tool_calls) {
      const args = JSON.parse(call.function.arguments);
      console.log(`[Agent] Executing tool: ${call.function.name}`, JSON.stringify(args));
      try {
        const result = await dispatchSuperDocTool(client, call.function.name, args);
        messages.push({ role: 'tool', tool_call_id: call.id, content: JSON.stringify(result) });
      } catch (error) {
        const errorMsg = error instanceof Error ? error.message : String(error);
        console.error(`[Agent] Tool error: ${errorMsg}`);
        messages.push({ role: 'tool', tool_call_id: call.id, content: JSON.stringify({ error: errorMsg }) });
      }
    }

    if (messages.length > 50) {
      console.log('[Agent] Too many iterations, stopping');
      break;
    }
  }

  return null;
}

/**
 * Wait for shutdown signal.
 */
async function waitForShutdown(): Promise<void> {
  console.log('[Agent] Listening for document changes. Press Ctrl+C to exit.');
  await new Promise((resolve) => {
    process.on('SIGINT', resolve);
    process.on('SIGTERM', resolve);
  });
  console.log('\n[Agent] Shutting down...');
}

/**
 * Main entry point.
 */
async function main() {
  const documentId = process.argv[2] || DEFAULT_DOC_ID;

  console.log('[Agent] Starting Document API agent...');
  console.log(`[Agent] Collaboration URL: ${COLLAB_URL}/${documentId}`);

  const client = await connectToSuperdoc();

  try {
    await joinCollaboration(client, documentId);

    // Initialize LLM tools
    const { tools, openai } = await initLLMTools();

    // Set up debounced creative writer
    const { trigger, cancel } = createDebouncedHandler(
      (signal) => readAndContinueWriting(client, openai, tools, signal),
      DEBOUNCE_MS,
    );

    // Connect to Yjs for real-time updates
    const { cleanup: cleanupYjs } = connectToYjsUpdates(documentId, () => {
      // Ignore updates from our own writes
      if (isAgentWriting) {
        console.log('[Agent] Ignoring own update...');
        return;
      }
      console.log('[Agent] Document changed, waiting for user to stop typing...');
      trigger();
    });

    await waitForShutdown();

    cancel();
    cleanupYjs();
    await client.doc.close({});
  } catch (error) {
    console.error('[Agent] Error:', error instanceof Error ? error.message : error);
    process.exit(1);
  } finally {
    await client.dispose();
    console.log('[Agent] Disconnected');
  }
}

main();

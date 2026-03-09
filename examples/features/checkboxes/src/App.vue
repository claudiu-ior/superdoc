<script setup lang="ts">
import { ref, onUnmounted, watch } from 'vue';
import { SuperDoc } from 'superdoc';
import 'superdoc/style.css';

/**
 * Checkboxes Example
 *
 * This example demonstrates how to:
 * 1. Add a custom "Insert Checkbox" button to the toolbar
 * 2. Insert checkbox content controls at the cursor position
 * 3. Toggle checkboxes programmatically
 * 4. List all checkboxes in the document
 */

const file = ref<File | null>(null);
const checkboxes = ref<any[]>([]);
const containerEl = ref<HTMLDivElement | null>(null);
let superdoc: any = null;

// Checkbox icon SVG
const checkboxIcon = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
  <rect x="3" y="3" width="18" height="18" rx="2" ry="2"/>
  <path d="M9 12l2 2 4-4"/>
</svg>`;

// Insert a checkbox at the current cursor position
const insertCheckbox = async () => {
  const editor = superdoc?.activeEditor;
  if (!editor) return;

  // Create a checkbox content control at the current cursor position
  const result = await editor.doc.create.contentControl({
    kind: 'inline',
    controlType: 'checkbox',
    tag: `checkbox-${Date.now()}`, // Unique tag for identification
  });

  if (!result.success) {
    console.error('Failed to create checkbox:', result.failure);
    return;
  }

  console.log('Checkbox inserted:', result.contentControl);

  // Move cursor to the right of the newly created checkbox
  const nodeId = result.contentControl.nodeId;
  let positionAfterCheckbox: number | null = null;

  editor.state.doc.descendants((node: any, pos: number) => {
    if (node.attrs?.id === nodeId) {
      positionAfterCheckbox = pos + node.nodeSize;
      return false; // Stop traversal
    }
  });

  if (positionAfterCheckbox !== null && editor.commands?.setTextSelection) {
    editor.commands.setTextSelection(positionAfterCheckbox);
    editor.commands.focus?.();
  }

  refreshCheckboxList();
};

// Toggle a specific checkbox using its target
const toggleCheckbox = async (checkbox: any) => {
  const editor = superdoc?.activeEditor;
  if (!editor) return;

  // Use the target property from ContentControlInfo
  const result = await editor.doc.contentControls.checkbox.toggle({
    target: checkbox.target,
  });

  if (result.success) {
    console.log('Checkbox toggled');
    refreshCheckboxList();
  } else {
    console.error('Failed to toggle checkbox:', result.failure);
  }
};

// Get the checked state of a checkbox
const getCheckboxState = (checkbox: any): boolean => {
  const editor = superdoc?.activeEditor;
  if (!editor) return false;

  try {
    // Use the target property from ContentControlInfo
    const state = editor.doc.contentControls.checkbox.getState({
      target: checkbox.target,
    });
    return state.checked;
  } catch (e) {
    console.error('Error getting checkbox state:', e);
    return false;
  }
};

// Refresh the list of checkboxes in the document
const refreshCheckboxList = () => {
  const editor = superdoc?.activeEditor;
  if (!editor) return;

  // List all content controls and filter for checkboxes
  const allControls = editor.doc.contentControls.list();
  console.log('All content controls:', allControls);

  const checkboxControls = allControls.items.filter(
    (item: any) => item.type === 'checkbox' || item.controlType === 'checkbox'
  );
  console.log('Checkbox controls:', checkboxControls);

  // Add checked state to each checkbox
  checkboxes.value = checkboxControls.map((cb: any) => ({
    ...cb,
    checked: getCheckboxState(cb),
  }));
};

// Handle file selection
const onFileChange = (event: Event) => {
  const target = event.target as HTMLInputElement;
  file.value = target.files?.[0] ?? null;
};

// Export the document to see checkbox in DOCX
const exportDocument = async () => {
  if (!superdoc) return;

  const blob = await superdoc.export({ format: 'docx' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'document-with-checkboxes.docx';
  a.click();
  URL.revokeObjectURL(url);
};

// Handle click on checkbox to toggle it
let lastCheckboxId: string | null = null;

const setupCheckboxClickHandler = () => {
  const editor = superdoc?.activeEditor;
  if (!editor) return;

  // Listen for selection changes - when user clicks on a checkbox, selection moves there
  editor.on('selectionUpdate', async ({ editor: ed }: any) => {
    // Get the current selection position
    const { state } = ed;
    const { $from } = state.selection;

    // Walk up the node tree to find if we're inside a structured content node
    let foundCheckboxId: string | null = null;

    for (let d = $from.depth; d >= 0; d--) {
      const node = $from.node(d);
      if (node.type.name === 'structuredContent' || node.type.name === 'structuredContentBlock') {
        const controlType = node.attrs?.controlType;

        if (controlType === 'checkbox') {
          foundCheckboxId = node.attrs?.id;
          break;
        }
      }
    }

    // If we're not in a checkbox, reset tracking
    if (!foundCheckboxId) {
      lastCheckboxId = null;
      return;
    }

    // Only toggle if this is a different checkbox than last time
    // (prevents toggling on every keystroke within the same checkbox)
    if (foundCheckboxId === lastCheckboxId) {
      return;
    }

    // Update tracking
    lastCheckboxId = foundCheckboxId;

    // Find and toggle the checkbox
    const allControls = ed.doc.contentControls.list();
    const checkbox = allControls.items.find(
      (item: any) => item.id === foundCheckboxId || item.target?.nodeId === foundCheckboxId
    );

    if (checkbox) {
      console.log('Toggling checkbox:', checkbox.id);
      await ed.doc.contentControls.checkbox.toggle({
        target: checkbox.target,
      });
      refreshCheckboxList();
    }
  });
};

// Watch for file changes and initialize SuperDoc
watch(file, async (newFile) => {
  if (!newFile || !containerEl.value) return;

  // Destroy previous instance
  superdoc?.destroy();

  // Create new SuperDoc instance
  superdoc = new SuperDoc({
    selector: containerEl.value,
    document: newFile,
    toolbar: '#toolbar',
    modules: {
      toolbar: {
        // Add the checkbox button to the toolbar
        customButtons: [
          {
            type: 'button',
            name: 'insertCheckbox',
            tooltip: 'Insert Checkbox',
            icon: checkboxIcon,
            group: 'center',
            command: insertCheckbox,
          },
        ],
      },
    },
    onReady: () => {
      // Load existing checkboxes when document is ready
      setTimeout(() => {
        refreshCheckboxList();
        setupCheckboxClickHandler();
      }, 500);
    },
  });
});

onUnmounted(() => {
  superdoc?.destroy();
  superdoc = null;
});
</script>

<template>
  <div class="app">
    <!-- Header -->
    <header class="header">
      <input type="file" accept=".docx" @change="onFileChange" />
      <button @click="insertCheckbox" :disabled="!file">Insert Checkbox</button>
      <button @click="refreshCheckboxList" :disabled="!file">Refresh List</button>
      <button @click="exportDocument" :disabled="!file">Export DOCX</button>
    </header>

    <!-- Toolbar -->
    <div id="toolbar"></div>

    <!-- Main content area -->
    <div class="main">
      <!-- Document editor -->
      <div ref="containerEl" class="editor"></div>

      <!-- Checkbox sidebar -->
      <aside v-if="file" class="sidebar">
        <h3>Checkboxes ({{ checkboxes.length }})</h3>
        <p v-if="checkboxes.length === 0" class="empty-message">
          No checkboxes found. Click "Insert Checkbox" to add one.
        </p>
        <ul v-else class="checkbox-list">
          <li v-for="(cb, index) in checkboxes" :key="cb.id" class="checkbox-item">
            <input
              type="checkbox"
              :checked="cb.checked"
              @change="toggleCheckbox(cb)"
            />
            <span class="checkbox-title">{{ cb.properties?.title || `Checkbox ${index + 1}` }}</span>
            <code class="checkbox-tag">{{ cb.properties?.tag }}</code>
          </li>
        </ul>
      </aside>
    </div>
  </div>
</template>

<style scoped>
.app {
  display: flex;
  flex-direction: column;
  height: 100vh;
}

.header {
  padding: 0.75rem 1rem;
  background: #f5f5f5;
  border-bottom: 1px solid #ddd;
  display: flex;
  gap: 1rem;
  align-items: center;
}

.main {
  display: flex;
  flex: 1;
  overflow: hidden;
}

.editor {
  flex: 1;
  overflow: auto;
}

.sidebar {
  width: 280px;
  border-left: 1px solid #ddd;
  padding: 1rem;
  overflow: auto;
  background: #fafafa;
}

.sidebar h3 {
  margin-bottom: 1rem;
}

.empty-message {
  color: #666;
  font-size: 0.9rem;
}

.checkbox-list {
  list-style: none;
  padding: 0;
}

.checkbox-item {
  padding: 0.5rem;
  margin-bottom: 0.5rem;
  background: #fff;
  border: 1px solid #ddd;
  border-radius: 4px;
  display: flex;
  align-items: center;
  gap: 0.5rem;
}

.checkbox-title {
  flex: 1;
}

.checkbox-tag {
  font-size: 0.7rem;
  color: #888;
}
</style>

<template>
    <div class="file-drop" @dragenter.prevent @dragover.prevent @drop="onDrop">
      <slot>
        <p>Drag and drop your .doc or .docx files here or click to select files</p>
      </slot>
    </div>
  </template>
  
  <script>
  export default {
    emits: ['files-dropped'],
    methods: {
      onDrop(event) {
        event.preventDefault(); // Add this line to prevent the default behavior
        const files = event.dataTransfer.files;
        this.$emit('files-dropped', files);
      },
    },
  };
  </script>
  
  <style scoped>
  .file-drop {
    display: flex;
    justify-content: center;
    align-items: center;
    width: 100%;
    height: 200px;
    border: 2px dashed #6b7280;
    border-radius: 4px;
    background-color: #f3f4f6;
    font-family: Arial, sans-serif;
    font-size: 16px;
    color: #6b7280;
    cursor: pointer;
  }
  
  .file-drop:hover {
    background-color: #e5e7eb;
  }
  </style>
  
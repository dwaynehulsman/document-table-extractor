<template>
    <div>
      <file-drop @files-dropped="handleFiles"></file-drop>
      <button :disabled="!hasData" @click="exportData">Download as XLS</button>
    </div>
  </template>
  
  <script>
  import FileDrop from './FileDrop.vue';
  import mammoth from '../mammoth-wrapper';
  import { utils, writeFile } from 'xlsx';
  
  
  export default {
    components: {
      FileDrop
    },
    data() {
      return {
        hasData: false,
        data: []
      };
    },
    methods: {
        async handleFiles(files) {
        const validFiles = Array.from(files).filter(file => file.name.endsWith('.doc') || file.name.endsWith('.docx'));

        for (const file of validFiles) {
            const data = await this.extractTables(file);
            // Flatten the array of arrays
            this.data.push(...data.flat());
        }

        this.hasData = this.data.length > 0;
        },

      extractTables(file) {
        return new Promise(resolve => {
          const reader = new FileReader();
          reader.onload = async event => {
            const arrayBuffer = event.target.result;
            const result = await mammoth.convertToHtml({ arrayBuffer });
            const tables = result.value.match(/<table[^>]*>([\s\S]*?)<\/table>/g);
            const tableData = tables.map(table => this.parseTable(table, file.name));
            resolve(tableData);
          };
          reader.readAsArrayBuffer(file);
        });
      },
      parseTable(htmlString, fileName) {
            const parser = new DOMParser();
            const dom = parser.parseFromString(htmlString, 'text/html');
            const rows = Array.from(dom.querySelectorAll('tr'));

            return rows.map((row, rowIndex) => {
            const rowData = Array.from(row.children).map(cell => cell.textContent.trim());
            rowData.push(fileName);
            
            if (rowIndex === 0 && this.isHeaderDuplicate(rowData)) {
                return null;
            }
            
            return rowData;
            }).filter(rowData => rowData !== null);
        },
        isHeaderDuplicate(headerRow) {
            if (this.data.length === 0) {
            return false;
            }
            
            const firstRow = this.data[0];
            const headerRowWithoutFilename = headerRow.slice(0, -1);
            const firstRowWithoutFilename = firstRow.slice(0, -1);

            return headerRowWithoutFilename.every((header, index) => header === firstRowWithoutFilename[index]);
        },
      exportData() {
        const wb = utils.book_new();
        const ws = utils.aoa_to_sheet(this.data);
        utils.book_append_sheet(wb, ws, 'Sheet1');
        writeFile(wb, 'exported_data.xls', { type: 'binary', bookType: 'xls' });
      }
    }
  };
  </script>
  
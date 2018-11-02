<template>
    <div>
        <file-upload
                ref="upload"
                :drop="true"
                :multiple="false"
                @input-file="inputFile"
        >
            <div class="dropzone">
                Drop XLSX file here or click to upload.
            </div>
        </file-upload>
        <div v-if="tableData.body.length > 0">
            <select v-model="database">
                <option value="mysql">MySQL</option>
                <option value="mssql">Microsoft SQL Server</option>
                <option value="oracle">Oracle</option>
            </select>
            <select v-model="queryMode">
                <option value="insert">INSERT Statements</option>
                <option value="merge_sqlserver" v-if="database === 'mssql'">MERGE Statements</option>
                <option value="replace" v-if="database==='mysql'">REPLACE Statements</option>
                <option value="select_oracle" v-if="database === 'oracle'">SELECT Query</option>
                <option value="select_mysql" v-if="database !== 'oracle'">SELECT Query</option>
            </select>
            <input v-model="tableName" v-if="database !== '' && (queryMode == 'insert' || queryMode == 'replace' || queryMode == 'merge_sqlserver')"><br>
            <textarea v-model="sqlQuery"></textarea>
        </div>
    </div>
</template>

<style scope>
    .dropzone {
        padding: 5rem 7rem;
        cursor: pointer;
        font-weight: bold;
        border-radius: 1rem;
        border: 3px dashed #aa99dd;
    }
    select {
        margin-top: 2rem;
    }
    textarea {
        width: 100%;
        min-height: 100px;
        margin-top: 1rem;
    }
</style>

<script>
  import XLSX from 'xlsx'
  import VueUploadComponent from 'vue-upload-component'

  export default {
    components: {
      FileUpload: VueUploadComponent
    },
    data: () => ({
      queryMode: 'insert',
      tableName: 'table_name',
      database: '',
      rawFile: null,
      workbook: null,
      tableData: {
        header: [],
        body: []
      }
    }),
    computed: {
      sqlQuery() {
        /*if (this.database === 'oracle') {
          let guard = (item) => ('\'' + (typeof(item) == 'string' ? item.replace(/'/g, "''").replace(/&/g, "'||chr(38)||'") : item) + '\'')
        }
        else {
          let guard = (item) => ('\'' + (typeof(item) == 'string' ? item.replace(/'/g, "''") : item) + '\'')
        }*/
        let guard = (item) => {
          if (typeof(item) == 'string') {
            item = item.replace(/'/g, "''")
            if (this.database === 'oracle') {
              item = item.replace(/&/g, "'||chr(38)||'")
            }
          }
          return `'${item}'`
        }
        if (this.tableData.body.length == 0) {
          return ''
        }
        var sqlQuery = ''

        if (this.queryMode == 'select_oracle' || this.queryMode == 'select_mysql') {
          var rowQuery = []

          this.tableData.body.forEach((row) => {
            rowQuery.push('SELECT ' +  Object.values(row).map((item, key) => (guard(item)+' as ' + this.tableData.header[key])).join(',') + (this.queryMode == 'select_oracle' ? ' FROM DUAL' : ''))
          })
          sqlQuery = rowQuery.join("\nUNION ALL\n")
        }
        if (this.queryMode == 'merge_sqlserver') {
          this.tableData.body.forEach((row) => {
              var values = Object.values(row).map((item) => guard(item) ).join(',')
              sqlQuery += ('MERGE INTO '+this.tableName+' as target USING (SELECT ' + values + ')' )
              sqlQuery += ' as source ('+ this.tableData.header.join(',')  +') on '
              sqlQuery += '(' + 'target.'+this.tableData.header[0]+' = source.'+this.tableData.header[0] +')'
              sqlQuery += ' WHEN MATCHED THEN UPDATE SET ' + this.tableData.header.map( (item,i) => {
                var rowItem = Object.values(row)[i]
                return item + ' = ' + guard(rowItem) } 
              )
              sqlQuery += ' WHEN NOT MATCHED THEN INSERT ('+this.tableData.header.join(',') +') VALUES ('+values +')'
              sqlQuery += ';\n'
            }
          )
         
        }
        if(this.queryMode == 'insert' || this.queryMode == 'replace') {
          var insertQuery = (this.queryMode == 'insert' ? 'INSERT' : 'REPLACE')  + ' INTO ' + this.tableName + '('+ this.tableData.header.join(',')  +')'
          this.tableData.body.forEach((row) => {
            sqlQuery += insertQuery + ' VALUES (' + Object.values(row).map((item) => guard(item)).join(',') + ');' + "\n"
          })
        }

        return sqlQuery
      }
    },
    methods: {
     
      inputFile: function (newFile, oldFile) {
        if (newFile) {
          this.rawFile = newFile.file
          this.convertToWorkbook()
        }
      },
      convertToWorkbook() {
        this.fileConvertToWorkbook(this.rawFile)
          .then((workbook) => {
            let xlsxArr = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]])
            this.workbook = workbook
            this.initTable(
              this.xlsxArrToTableArr(xlsxArr)
            )
            if (this.workbook !== null) {
              this.tableName = this.workbook.SheetNames[0]
            }
          })
          .catch((err) => {
            console.error(err)
          })
      },
      fileConvertToWorkbook(file) {
        let reader = new FileReader()
        let fixdata = (data) => {
          let o = "", l = 0, w = 10240
          for (; l < data.byteLength / w; ++l) {
            o += String.fromCharCode.apply(null, new Uint8Array(data.slice(l * w, l * w + w)))
          }
          o += String.fromCharCode.apply(null, new Uint8Array(data.slice(l * w)))
          return o
        }
        return new Promise((resolve, reject) => {
          try {
            reader.onload = (renderEvent) => {
              let data = renderEvent.target.result
              if (this.rABS) {
                /* if binary string, read with type 'binary' */
                resolve(XLSX.read(data, {type: 'binary'}))
              } else {
                /* if array buffer, convert to base64 */
                let arr = fixdata(data)
                resolve(XLSX.read(btoa(arr), {type: 'base64'}))
              }
            }
            reader.onerror = (error) => {
              reject(error)
            }
            if (this.rABS) {
              reader.readAsBinaryString(file)
            } else {
              reader.readAsArrayBuffer(file)
            }
          } catch (error) {
            reject(error)
          }
        })
      },
      xlsxArrToTableArr(xlsxArr) {
        let tableArr = []
        let length = 0
        let maxLength = 0
        let maxLengthIndex = 0
        xlsxArr.forEach((item, index) => {
          length = Object.keys(item).length
          if (maxLength < length) {
            maxLength = length
            maxLengthIndex = index
          }
        })
        let tableHeader = Object.keys(xlsxArr[maxLengthIndex])
        let rowItem = {}
        xlsxArr.forEach((item) => {
          rowItem = {}
          for (let i = 0; i < maxLength; i++) {
            rowItem[tableHeader[i]] = item[tableHeader[i]] || ''
          }
          tableArr.push(rowItem)
        })
        return {
          header: tableHeader,
          data: tableArr
        }
      },
      tableArrToXlsxArr({data, header}) {
        let xlsxArr = []
        let tempObj = {}
        data.forEach((rowItem) => {
          tempObj = {}
          rowItem.forEach((item, index) => {
            tempObj[header[index]] = item
          })
          xlsxArr.push(tempObj)
        })
        return xlsxArr
      },
      initTable({data, header}) {
        this.tableData.header = header
        this.tableData.body = data
      }
    }
  }
</script>
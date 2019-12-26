<template>
    <div>
        Upload file: <input type="file" @change="handleChange" /> {{ file_info }}
        <table>
            <tr v-for="(row,ri) in this.data_list" v-bind:key="ri">
                <td v-for="(cell, ci) in row" :key="ci">
                    {{cell}}
                </td>
            </tr>
        </table>
        <div>
            <div v-for="(r, i) in this.csv_list" :key="i" class="csv">
                {{ r }}
            </div>
        </div>
    </div>
</template>

<script>
    // Используем exceljs так как excel4node не умеет читать, только создавать
    /*
        // i use the vue template
        <input type="file" @change="handleChange" />

        handleChange(e) {
          this.file = e.target.files[0]
        },
        handleImport() {
          const wb = new Excel.Workbook();
          const reader = new FileReader()

          reader.readAsArrayBuffer(this.file)
          reader.onload = () => {
            const buffer = reader.result;
            wb.xlsx.load(buffer).then(workbook => {
              console.log(workbook, 'workbook instance')
              workbook.eachSheet((sheet, id) => {
                sheet.eachRow((row, rowIndex) => {
                  console.log(row.values, rowIndex)
                })
              })
            })
          }
        }
    */



    import Excel from 'exceljs'
    export default {
        name: "XlsUpload",
        data(){
            return{
                file: null,
                file_info: '',
                data_list:[],
                csv_list:[],
            }
        },
        watch:{
            file(){
                this.file_info ='type: '+this.file.type+' size: '+this.file.size+' byte';
                /* eslint-disable */
                // console.log('name: '+this.file.name)
                // console.log('size: '+this.file.size)
                console.log('type: '+this.file.type)
                /* eslint-enable */
                if (this.file.type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'){
                    this.ExlsImport();
                }
                else if (this.file.type === 'text/csv'){
                    this.CsvImport();
                }
            }
        },
        methods:{
            handleChange(e) {
                this.file = e.target.files[0]
            },
            CsvImport(){
                /* eslint-disable */
                console.log('CsvImport')

                let reader = new FileReader();

                reader.readAsText(this.file, 'windows-1251');

                reader.onload = () => {
                    let textByLine = reader.result.split("\n")
                    textByLine.forEach(row => {
                        row = row.trim();
                        if (row.length>0){
                            this.csv_list.push(row);
                        }
                        //console.log(row)
                    });
                    //console.log(reader.result);
                    console.log('onload');
                };

                reader.onerror = function() {
                    console.log(reader.error);
                };


                //let textByLine = this.file.split("\n")
                /* eslint-enable */
            },
            ExlsImport() {
                const wb = new Excel.Workbook();
                const reader = new FileReader()
                /* eslint-disable */
                reader.readAsArrayBuffer(this.file)
                reader.onload = () => {
                    const buffer = reader.result;
                    wb.xlsx.load(buffer).then(workbook => {
                        console.log(workbook, 'workbook instance')
                        workbook.eachSheet((sheet, id) => {
                            sheet.eachRow((row, rowIndex) => {
                                this.data_list.push(row.values)
                                console.log(rowIndex, row.values)
                            })
                        })
                    })
                }
                /* eslint-enable */
            }
        }
    }
</script>

<style scoped>
    td, div.csv{
        padding: 1px;
        margin: 1px;
        border: solid 1px #8c8c8c;
    }
</style>
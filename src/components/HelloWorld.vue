<template>
    <div class="hello">
        <h3>{{ msg }}
            <a href="https://www.npmjs.com/package/exceljs" target="_blank">ExcelJs Npm</a>
            <a href="https://github.com/exceljs/exceljs" target="_blank">ExcelJs Github</a>
        </h3>
        <p>
            <button @click="getExcel">Get excel</button>
            <button @click="getCsv">Get CSV</button>
            <button @click="getTxtFile">Get txt</button>
        </p>

    </div>
</template>

<script>
    // https://www.npmjs.com/package/exceljs
    // https://github.com/exceljs/exceljs
    import Excel from 'exceljs'
    let workbook = new Excel.Workbook();
    let worksheet = workbook.addWorksheet('My Sheet');
    worksheet.columns = [
        { header: 'Id', key: 'id', width: 10 },
        { header: 'Column 1', key: 'k1', width: 32 },
        { header: 'Column 2', key: 'k2', width: 32 },
        { header: 'Column 3', key: 'k3', width: 32 },
        { header: 'Column 4', key: 'k4', width: 32 },
        { header: 'Name', key: 'name', width: 32 }
    ];

    worksheet.getColumn(4).outlineLevel = 0;
    worksheet.getColumn(5).outlineLevel = 1;

    worksheet.addRow({id: 1, name: 'John One'});
    worksheet.addRow({id: 2, name: 'John Two'});

    worksheet.addRow({k1: 1, k2: '1 One'});
    worksheet.addRow({k3: 2, k4: '2 Two'});

    worksheet.addRow([new Date('2019-08-05'), 5, 'Mid'], 1);

    export default {
        name: 'HelloWorld',
        props: {
            msg: String
        },
        methods:{
            getExcel(){
                // eslint-disable-next-line no-console
                console.log("!");
                // window.URL.createObjectURL(bb)
                workbook.xlsx.writeBuffer()
                    .then(function(buffer) {
                        // eslint-disable-next-line no-console
                        console.log("!!");
                        let link = document.createElement('a');
                        link.download = 'file.xlsx';
                        //let blob = new Blob([buffer], {type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
                        let blob = new Blob([buffer], {type: 'application/vnd.ms-excel'});
                        link.href = URL.createObjectURL(blob);
                        link.click();
                        URL.revokeObjectURL(link.href);
                    });
            },
            getCsv(){
                workbook.csv.writeBuffer()
                    .then(function(buffer) {
                        // eslint-disable-next-line no-console
                        console.log(buffer);
                        let link = document.createElement('a');
                        link.download = 'file.csv';
                        let blob = new Blob([buffer], {type: 'text/plain'});
                        link.href = URL.createObjectURL(blob);
                        link.click();
                        URL.revokeObjectURL(link.href);
                    });
            },
            getTxtFile(){
                let link = document.createElement('a');
                link.download = 'hello.txt';
                let blob = new Blob(['hello!'], {type: 'text/plain'});
                link.href = URL.createObjectURL(blob);
                link.click();
                URL.revokeObjectURL(link.href);
            }
        }
    }
</script>

<!-- Add "scoped" attribute to limit CSS to this component only -->

<style scoped>
    h3 {
        margin: 40px 0 0;
    }
    ul {
        list-style-type: none;
        padding: 0;
    }
    li {
        display: inline-block;
        margin: 0 10px;
    }

    .hello{
        text-align: center;
    }

</style>

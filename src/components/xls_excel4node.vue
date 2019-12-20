<template>
    <div>
        <a href="https://www.npmjs.com/package/excel4node" target="_blank">excel4node Npm</a>
        <button @click="getFile">Get Excel file</button>
    </div>
</template>

<script>
    // https://www.npmjs.com/package/excel4node
    import xl from 'excel4node';
    let wb = new xl.Workbook();
    let ws = wb.addWorksheet('Страница 1');
    ws.cell(1, 1).string('My simple string');
    ws.cell(1, 2).number(5);
    ws.cell(1, 3).formula('B1 * 10');
    ws.cell(1, 4).date(new Date());
    ws.cell(1, 5).link('http://parpar.in');
    ws.cell(1, 6).bool(true);

    ws.cell(3, 1, 3, 6).number(1); // All 6 cells set to number 1
    ws.cell(4, 2, 5, 5, true).string('One big merged cell');

    export default {
        name: "xls_excel4node",
        methods:{
            getFile(){
                // eslint-disable-next-line no-console
                console.log('!')
                wb.writeToBuffer().then(function(buffer) {
                    let link = document.createElement('a');
                    link.download = 'file.xlsx';
                    //let blob = new Blob([buffer], {type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
                    let blob = new Blob([buffer], {type: 'application/vnd.ms-excel'});
                    link.href = URL.createObjectURL(blob);
                    link.click();
                    URL.revokeObjectURL(link.href);
                });
            }
        }
    }
</script>

<style scoped>

</style>
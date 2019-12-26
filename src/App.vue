<template>
    <div id="app">
        <!--    <img alt="Vue logo" src="./assets/logo.png">-->
        <!--      https://www.npmjs.com/package/exceljs     -->
        <!--      https://github.com/exceljs/exceljs        -->
        <div v-if="false">
            <HelloWorld msg="Эксперементы с exceljs"/>
            <Xls_excel4node></Xls_excel4node>
        </div>

        <XlsUpload></XlsUpload>

    </div>
</template>

<script>
    import HelloWorld from './components/HelloWorld.vue'
    import Xls_excel4node from "./components/xls_excel4node";
    import XlsUpload from "./components/XlsUpload";

    export default {
        name: 'app',
        components: {
            HelloWorld,
            Xls_excel4node,
            XlsUpload
        }
    }
</script>

<style>
    #app {
        font-family: 'Avenir', Helvetica, Arial, sans-serif;
        -webkit-font-smoothing: antialiased;
        -moz-osx-font-smoothing: grayscale;
        text-align: center;
        color: #2c3e50;
        margin-top: 60px;
    }
    button{
        margin: 16px;
        padding: 6px;
    }
    a {
        color: #42b983;
        margin: 0 16px;
    }
</style>

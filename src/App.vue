<template>
  <div id="app">
    <h1>xlsx2sql - the easy way to get SQL queries from your xlsx files</h1>
    <p>Just upload a file below and you'll get your SQL queries in no time. Everything happens in your browser, no data is sent to the server!</p>

    <xlsx-converter/>
  </div>
</template>

<script>
import XlsxConverter from './components/XlsxConverter'

export default {
  name: 'App',
  components: {
    XlsxConverter
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
  max-width: 750px;
  margin-left: auto;
  margin-right: auto;
  background-color: #fff;
  -moz-box-shadow: 0 3px 5px #aaaaaa;
  -webkit-box-shadow: 0 3px 5px #aaaaaa;
  box-shadow: 0 3px 5px #aaaaaa;
  -moz-border-radius: 5px;
  border-radius: 5px;
  padding: 3rem;
}

  h1 {
    margin: 0 0 3rem 0;
  }
  p {
    margin: 0 0 3rem 0;
  }
</style>

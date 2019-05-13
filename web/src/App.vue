<template>
  <div id="app">
    <vue-dropzone
      v-show="loaded"
      @vdropzone-success="onSuccess"
      ref="myVueDropzone"
      id="dropzone"
      :options="dropzoneOptions"
    ></vue-dropzone>
    <table v-show="!loaded">
      <tr>
        <th>業務員</th>
        <th>客戶簡稱</th>
        <th>貨品名稱</th>
        <th>銷貨金額</th>
      </tr>
      <!-- <tr v-for="(item,i) in computedArray" :key="i">
        <td>{{item.employee}}</td>
        <td>{{item.customer}}</td>
        <td>{{item.product}}</td>
        <td>{{item.revenue}}</td>
      </tr>-->
    </table>
    <div v-for="(value,key,index) in computedArray" :key="index">
      {{key}}
      <div v-for="(value2,key2,index2) in computedArray[key]" :key="index2">
        {{key2}}
        <div v-for="(value3,key3,index3) in computedArray[key][key2]" :key="index3">
          {{key3}}
          <div>{{value3}}</div>
        </div>
      </div>
    </div>
  </div>
</template>

<script>
/* eslint-disable */
import vue2Dropzone from "vue2-dropzone";
import "vue2-dropzone/dist/vue2Dropzone.min.css";
import XLSX from "xlsx";
export default {
  name: "app",
  components: {
    vueDropzone: vue2Dropzone
  },
  data() {
    return {
      dropzoneOptions: {
        url: "https://httpbin.org/post",
        thumbnailWidth: 150,
        maxFilesize: 0.5
      },
      commissionedItems: ["硬化劑", "發泡劑", "設備", "薄膜"],
      array: "",
      loaded: true
    };
  },
  computed: {
    computedArray() {
      if (this.array) {
        var toReturn = this.array
          .filter(ele => {
            return this.commissionedItems.some(item =>
              ele["貨品名稱"].includes(item)
            );
          })
          .map(ele => {
            return {
              employee: ele["業務員"],
              customer: ele["客戶簡稱"],
              product: ele["貨品名稱"],
              revenue: ele["銷貨金額"]
            };
          })
          .sort((a, b) => {
            let empA = a.employee.toUpperCase();
            let empB = b.employee.toUpperCase();
            let comparison = 0;
            if (empA > empB) {
              comparison = 1;
            } else if (empA < empB) {
              comparison = -1;
            }
            return comparison;
          });

        var result = {};
        while (toReturn.length) {
          var object = toReturn.shift();
          // if employee exists
          if (result[object.employee]) {
            // if customer exists
            if (result[object.employee][object.customer]) {
              // if product exists
              if (result[object.employee][object.customer][object.product]) {
                // add revenue
                result[object.employee][object.customer][object.product] +=
                  object.revenue;
              } else {
                result[object.employee][object.customer][object.product] =
                  object.revenue;
              }
            } else {
              result[object.employee][object.customer] = {
                [object.product]: object.revenue
              };
            }
          } else {
            result[object.employee] = {
              [object.customer]: {
                [object.product]: object.revenue
              }
            };
          }
        }

        return result;
      }
      return "";
    }
  },
  methods: {
    onSuccess(file, response) {
      var f = file;
      var reader = new FileReader();
      var vm = this;
      reader.onload = function(e) {
        var data = new Uint8Array(e.target.result);
        var workbook = XLSX.read(data, { type: "array" });
        var sheet1 = workbook.SheetNames[0];
        var worksheet = workbook.Sheets[sheet1];
        vm.array = XLSX.utils.sheet_to_json(worksheet);
        vm.loaded = false;
      };
      reader.readAsArrayBuffer(f);
    }
  }
};
</script>

<style>
#app {
  font-family: "Avenir", Helvetica, Arial, sans-serif;
  -webkit-font-smoothing: antialiased;
  -moz-osx-font-smoothing: grayscale;
  text-align: center;
  color: #2c3e50;
}
* {
  margin: 0;
}
#dropzone {
  height: 100vh;
}
.dropzone .dz-message {
  text-align: center;
  margin: 50vh;
}
</style>

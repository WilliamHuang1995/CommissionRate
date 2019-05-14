<template>
  <div id="app">
    <h1 v-show="loaded">Please upload only Excel File, make sure raw data is on the first sheet.</h1>
    <vue-dropzone
      v-show="loaded"
      @vdropzone-success="onSuccess"
      ref="myVueDropzone"
      id="dropzone"
      :options="dropzoneOptions"
    ></vue-dropzone>
    <!-- <json-viewer v-show="!loaded" :value="computedArray" :expand-depth="4" copyable></json-viewer> -->
    <table v-if="array">
      <tbody>
        <tr>
          <th>業務員</th>
          <th>客戶簡稱</th>
          <th>貨品名稱</th>
          <th>銷貨金額</th>
          <th>0.5%</th>
          <th>1.5%</th>
          <th>3.0%</th>
          <th>5.0%</th>
        </tr>
        <template v-for="(employees,ekey) in computedArray">
          <template v-for="(customer,ckey) in computedArray[ekey]">
            <template v-for="(product,pkey,i) in computedArray[ekey][ckey]">
              <tr :key="i">
                <td>{{ekey}}</td>
                <td>{{ckey}}</td>
                <td>{{pkey}}</td>
                <td>{{product['0.5%'].toFixed(2)}}</td>
                <td>{{product['1.5%'].toFixed(2)}}</td>
                <td>{{product['3.0%'].toFixed(2)}}</td>
                <td>{{product['3.5%'].toFixed(2)}}</td>
                <td>{{product['5.0%'].toFixed(2)}}</td>
              </tr>
            </template>
          </template>
        </template>
      </tbody>
    </table>
  </div>
  <!-- use this https://tonygermaneri.github.io/canvas-datagrid/tutorials/vueExample.html -->
</template>

<script>
/* eslint-disable */
import vue2Dropzone from "vue2-dropzone";
import "vue2-dropzone/dist/vue2Dropzone.min.css";
import XLSX from "xlsx";
import JsonViewer from "vue-json-viewer";
import "vue-json-viewer/style.css";
export default {
  name: "app",
  components: {
    vueDropzone: vue2Dropzone,
    JsonViewer
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
                var revenue =
                  result[object.employee][object.customer][object.product][
                    "原本"
                  ] + object.revenue;
                result[object.employee][object.customer][object.product] = {
                  原本: revenue,
                  "0.5%": revenue * 0.005,
                  "1.5%": revenue * 0.015,
                  "3.0%": revenue * 0.03,
                  "3.5%": revenue * 0.035,
                  "5.0%": revenue * 0.05
                };
              } else {
                result[object.employee][object.customer][object.product] = {
                  原本: object.revenue,
                  "0.5%": object.revenue * 0.005,
                  "1.5%": object.revenue * 0.015,
                  "3.0%": object.revenue * 0.03,
                  "3.5%": object.revenue * 0.035,
                  "5.0%": object.revenue * 0.05
                };
              }
            } else {
              result[object.employee][object.customer] = {
                [object.product]: {
                  原本: object.revenue,
                  "0.5%": object.revenue * 0.005,
                  "1.5%": object.revenue * 0.015,
                  "3.0%": object.revenue * 0.03,
                  "3.5%": object.revenue * 0.035,
                  "5.0%": object.revenue * 0.05
                }
              };
            }
          } else {
            result[object.employee] = {
              [object.customer]: {
                [object.product]: {
                  原本: object.revenue,
                  "0.5%": object.revenue * 0.005,
                  "1.5%": object.revenue * 0.015,
                  "3.0%": object.revenue * 0.03,
                  "3.5%": object.revenue * 0.035,
                  "5.0%": object.revenue * 0.05
                }
              }
            };
          }
        }
        var final = [];
        return result;
      }
      return "";
    },
    computedArray2() {
      if (this.array) {
        return this.array
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
        console.log(vm.array);
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
  color: #2c3e50;
}
* {
  margin: 0;
  box-sizing: border-box;
}
#dropzone {
  height: 100vh;
}
.dropzone .dz-message {
  text-align: center;
  margin: 50vh;
}
table {
  width: 100%;
  text-align: center;
}
</style>

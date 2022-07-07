<template>
  <div>
    <DxBox :height=100 direction="col" width="100%">
      <!-- <DxBox :height="50" direction="col" width="100%"> -->
        <DxItem :ratio="1">
          <template #default>
            <div class="rect demo-dark header">
              Hasil Monitoring BMKGSoft Bulan Mei 2022
            </div>
          </template>
        </DxItem>
      <!-- </DxBox> -->

      <br />

        <DxItem :ratio="1">
          <template #default>
            <div class="rect demo-dark header">Balai Besar MKG Wilayah 1</div>
          </template>
        </DxItem>
      <br />
      <!-- <DxGroupPanel :visible="true"  name="City" /> -->
      <!-- <DxGrouping              data-field="City"
 :auto-expand-all="true" /> -->
 <DxDataGrid
          id="gridContainer"
          :data-source="employees"
          key-expr="ID"
          :show-borders="true"
          @exporting="onExporting"
        >
      <DxItem>
        
          <DxExport :enabled="true" :allow-export-selected-data="true" />

          <DxColumn data-field="Prefix" :width="40" />
          <DxColumn data-field="Stasiun" :width="200" />
          <DxColumn data-field="1" :width="40" />
          <DxColumn data-field="2" :width="40" />
          <DxColumn data-field="3" :width="40" />
          <DxColumn data-field="4" :width="40" />
          <DxColumn data-field="5" :width="40" />
          <DxColumn data-field="6" :width="40" />
          <DxColumn data-field="7" :width="40" />
          <DxColumn data-field="8" :width="40" />
          <DxColumn data-field="9" :width="40" />
          <DxColumn data-field="10" :width="40" />

          <DxColumn data-field="11" :width="40" />
          <DxColumn data-field="12" :width="40" />
          <DxColumn data-field="13" :width="40" />
          <DxColumn data-field="14" :width="40" />
          <DxColumn data-field="15" :width="40" />
          <DxColumn data-field="16" :width="40" />
          <DxColumn data-field="17" :width="40" />
          <DxColumn data-field="18" :width="40" />
          <DxColumn data-field="19" :width="40" />
          <DxColumn data-field="20" :width="40" />

          <DxColumn data-field="21" :width="40" />
          <DxColumn data-field="22" :width="40" />
          <DxColumn data-field="23" :width="40" />
          <DxColumn data-field="24" :width="40" />
          <DxColumn data-field="25" :width="40" />
          <DxColumn data-field="26" :width="40" />
          <DxColumn data-field="27" :width="40" />
          <DxColumn data-field="28" :width="40" />
          <DxColumn data-field="29" :width="40" />
          <DxColumn data-field="30" :width="40" />
          <DxColumn data-field="31" :width="40" />
      </DxItem>
              </DxDataGrid>

      <DxItem :ratio="1">
        <template #default>
          <div class="rect demo-dark footer">Footer</div>
        </template>
      </DxItem>
    </DxBox>
  </div>
</template>
<script>
import {
  DxDataGrid,
  DxColumn,
  DxExport,
  // DxItem,DxBox
  // DxGroupPanel,
  // DxGrouping,DxItem,
} from "devextreme-vue/data-grid";
import { DxBox, DxItem } from "devextreme-vue/box";

import { Workbook } from "exceljs";
import { saveAs } from "file-saver";
// Our demo infrastructure requires us to use 'file-saver-es'.
// We recommend that you use the official 'file-saver' package in your applications.
import { exportDataGrid } from "devextreme/excel_exporter";
import service from "./data.js";

export default {
  components: {
    DxDataGrid,
    DxColumn,
    DxExport,
    DxItem,
    DxBox,

    // DxGroupPanel,
    // DxGrouping,DxItem
  },
  data() {
    return {
      employees: service.getEmployees(),
    };
  },
  methods: {
    onExporting(e) {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet("Employees");

      exportDataGrid({
        component: e.component,
        worksheet,
        autoFilterEnabled: true,
      }).then(() => {
        workbook.xlsx.writeBuffer().then((buffer) => {
          saveAs(
            new Blob([buffer], { type: "application/octet-stream" }),
            "DataGrid.xlsx"
          );
        });
      });
      e.cancel = true;
    },
  },
};
</script>

<!-- <style scoped>
#gridContainer {
  height: 423px;
}
</style> -->
<style>
.rect {
  text-align: center;
  font-size: 25px;
  font-weight: bold;
  padding-top: 10px;
  height: 100%;
}

.demo-light {
  background: rgba(245, 229, 166, 0.5);
}

.demo-dark {
  background: rgba(148, 215, 199, 0.5);
}

.demo-dark.header {
  background: rgba(243, 158, 108, 0.5);
}

.demo-dark.footer {
  background: rgba(123, 155, 207, 0.5);
}

.small {
  height: 50px;
  border: 1px solid lightgray;
}
</style>

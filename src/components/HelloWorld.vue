<template>
  <div>
    <DxBox :height="100" direction="col">
      <!-- <DxBox :height="50" direction="col" width="100%"> -->
      <DxItem :ratio="1">
        <template #default>
          <div class="rect demo-dark header">
            Hasil Monitoring BMKGSoft Bulan Mei 2022
          </div>
        </template>
      </DxItem>

      <br />

      <DxItem :ratio="1">
        <template #default>
          <div class="rect demo-dark header">Balai Besar MKG Wilayah 1</div>
        </template>
      </DxItem>
    </DxBox>
    <br />

    <DxBox :height="30" direction="col">
      <DxItem :ratio="1">
        <template #default>
          <div class="tipe demo-dark header">Tipe FORM: Me-48</div>
        </template>
      </DxItem>
      <br />
    </DxBox>
    <!-- <DxGroupPanel :visible="true"  name="City" /> -->
    <!-- <DxGrouping              data-field="City"
 :auto-expand-all="true" /> -->
    <DxDataGrid
      id="gridContainer"
      :data-source="employees"
      key-expr="ID"
      :word-wrap-enabled="true"
      :show-borders="true"
      :alignment="center"
      @exporting="onExporting"
      CssClass="customGrid"
    >
      <!-- <DxItem> -->

      <DxExport :enabled="true" :allow-export-selected-data="true" />

      <DxColumn
        data-field="Prefix"
        caption="No"
        vertical-alignment="middle"
        :width="40"
        alignment=" center"
      />
      <!-- :fixed="true" -->

      <DxColumn data-field="Stasiun" :width="200" alignment="center" />
      <DxColumn caption="Tanggal" alignment="center">
        <DxColumn data-field="1" :width="40" alignment="center" />
        <DxColumn data-field="2" :width="40" alignment="center" />
        <DxColumn data-field="3" :width="40" alignment="center" />
        <DxColumn data-field="4" :width="40" alignment="center" />
        <DxColumn data-field="5" :width="40" alignment="center" />
        <DxColumn data-field="6" :width="40" alignment="center" />
        <DxColumn data-field="7" :width="40" alignment="center" />
        <DxColumn data-field="8" :width="40" alignment="center" />
        <DxColumn data-field="9" :width="40" alignment="center" />
        <DxColumn data-field="10" :width="40" alignment="center" />

        <DxColumn data-field="11" :width="40" alignment="center" />
        <DxColumn data-field="12" :width="40" alignment="center" />
        <DxColumn data-field="13" :width="40" alignment="center" />
        <DxColumn data-field="14" :width="40" alignment="center" />
        <DxColumn data-field="15" :width="40" alignment="center" />
        <DxColumn data-field="16" :width="40" alignment="center" />
        <DxColumn data-field="17" :width="40" alignment="center" />
        <DxColumn data-field="18" :width="40" alignment="center" />
        <DxColumn data-field="19" :width="40" alignment="center" />
        <DxColumn data-field="20" :width="40" alignment="center" />

        <DxColumn data-field="21" :width="40" alignment="center" />
        <DxColumn data-field="22" :width="40" alignment="center" />
        <DxColumn data-field="23" :width="40" alignment="center" />
        <DxColumn data-field="24" :width="40" alignment="center" />
        <DxColumn data-field="25" :width="40" alignment="center" />
        <DxColumn data-field="26" :width="40" alignment="center" />
        <DxColumn data-field="27" :width="40" alignment="center" />
        <DxColumn data-field="28" :width="40" alignment="center" />
        <DxColumn data-field="29" :width="40" alignment="center" />
        <DxColumn data-field="30" :width="40" alignment="center" />
        <DxColumn data-field="31" :width="40" alignment="center" />
      </DxColumn>
      <DxColumn
        caption="Rata-rata Data Masuk"
        alignment="center"
        :width="70"
        :height="100"
      />
      <DxColumn
        :width="90"
        caption="Penyebab Data Tidak Masuk"
        alignment="center"
      />

      <!-- </DxItem> -->
    </DxDataGrid>

    <DxItem :ratio="1">
      <template #default>
        <div class="rect demo-dark footer">Footer</div>
      </template>
    </DxItem>

    <!-- </DxBox> -->
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
        topLeftCell: { row: 8, column: 1 },

        autoFilterEnabled: true,
      })
        .then((cellRange) => {
          // header
          const headerRowHasil = worksheet.getRow(2);
          headerRowHasil.height = 30;
          worksheet.mergeCells(2, 1, 2, 35);

          headerRowHasil.getCell(1).value =
            "Hasil Monitoring BMKGSoft Bulan Mei 2022";
          headerRowHasil.getCell(1).font = {
            name: "Dyuthi",
            size: 22,
            fontWeight: "bold",
            color: "#2acaea",
          };
          headerRowHasil.getCell(1).alignment = {
            horizontal: "center",
          };

          const headerRowBalai = worksheet.getRow(4);
          headerRowBalai.height = 30;
          worksheet.mergeCells(4, 1, 4, 35);

          headerRowBalai.getCell(1).value = "Balai Besar MKG Wilayah 1";

          headerRowBalai.getCell(1).font = {
            name: "Dyuthi",
            size: 12,
            verticalAlignment: "center ",
            fontWeight: "bold",
            color: { argb: "FF0000FF" },
          };
          headerRowBalai.getCell(1).alignment = {
            horizontal: "center",
            vertical: "middle",
          };

          const headerRowTipe = worksheet.getRow(6);
          headerRowTipe.height = 20;
          worksheet.mergeCells(6, 1, 6, 7);

          headerRowTipe.getCell(1).value = "Tipe FORM: Me-48";
          headerRowTipe.getCell(1).font = { name: "Segoe UI Light", size: 14 };
          headerRowTipe.getCell(1).alignment = { horizontal: "left" };

          // footer
          const footerRowIndex = cellRange.to.row + 3;
          const footerRow = worksheet.getRow(footerRowIndex);
          worksheet.mergeCells(footerRowIndex, 1, footerRowIndex, 4);

          footerRow.getCell(1).value = "Mengetahui.";
          footerRow.getCell(1).font = {
            color: { argb: "#000" },
            // italic: true,
          };
          footerRow.getCell(1).alignment = { horizontal: "left" };

          // Sub
          const footerRowIndexKoordinator = cellRange.to.row + 4;
          const footerRowKoordinator = worksheet.getRow(
            footerRowIndexKoordinator
          );
          worksheet.mergeCells(
            footerRowIndexKoordinator,
            1,
            footerRowIndexKoordinator,
            4
          );

          footerRowKoordinator.getCell(3).value = "Sub Koordinator";
          footerRowKoordinator.getCell(3).font = {
            color: { argb: "#000" },
            italic: true,
          };
          footerRow.getCell(1).alignment = { horizontal: "left" };

          // Balai
          const footerRowIndexBidang = cellRange.to.row + 5;
          const footerRowBidang = worksheet.getRow(footerRowIndexBidang);
          worksheet.mergeCells(
            footerRowIndexBidang,
            1,
            footerRowIndexBidang,
            4
          );

          footerRowBidang.getCell(3).value = "Bidang Manajemen Database MKG,";
          footerRowBidang.getCell(3).font = {
            color: "red",
            style: "bold",
            // name: "Times-Bold",
            // italic: true,
          };
          footerRow.getCell(1).alignment = { horizontal: "left" };

          // Jakarta
          const footerRowIndexJakarta = cellRange.to.row + 4;
          const footerRowJakarta = worksheet.getRow(footerRowIndexJakarta);
          worksheet.mergeCells(
            footerRowIndexJakarta,
            29,
            footerRowIndexJakarta,
            33
          );
          // worksheet.mergeCells(
          //   footerRowIndexBidang,
          //   29,
          //   footerRowIndexBidang,
          //   33
          // );
          footerRowJakarta.getCell(29).value = "Jakarta,";
          footerRowJakarta.getCell(4).font = {
            color: { argb: "#000" },
            // italic: true,
          };
          footerRow.getCell(4).alignment = { horizontal: "left" };

          // Pembuat Laporan,
          const footerRowIndexPembuat = cellRange.to.row + 5;
          const footerRowPembuat = worksheet.getRow(footerRowIndexPembuat);
          // worksheet.mergeCells(
          //   footerRowIndexJakarta,
          //   29,
          //   footerRowIndexJakarta,
          //   33
          // );
          worksheet.mergeCells(
            footerRowIndexPembuat,
            29,
            footerRowIndexPembuat,
            33
          );
          footerRowPembuat.getCell(29).value = "Pembuat Laporan,";
          footerRowPembuat.getCell(4).font = {
            color: { argb: "#000" },
            // italic: true,
          };
          footerRow.getCell(4).alignment = { horizontal: "left" };
        })
        .then(() => {
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
.dx-datagrid-headers.dx-header-multi-row
  .dx-datagrid-content
  .dx-datagrid-table
  .dx-row.dx-header-row
  > td {
  vertical-align: middle;
  text-align: center vertically;
  font-weight: bold;
}
.rect {
  text-align: center;
  vertical-align: middle;
  font-size: 25px;
  font-weight: bold;
  padding-top: 10px;
  height: 100%;
  color: #2acaea;
}

.tipe {
  text-align: left;
  font-size: 16px;
  font-weight: bold;
  padding-top: 10px;
  height: 100%;
}

/* .demo-light {
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
} */
</style>

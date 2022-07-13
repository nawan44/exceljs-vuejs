<template>
  <div id="app">
    <img alt="Vue logo" src="./assets/logo.png" width="25%" />
    <button @click="exportExcel">Export ExcelJs</button>

    <HelloWorld />
  </div>
</template>

<script>
import ExcelJS from "exceljs";
import { saveAs } from "file-saver";
import HelloWorld from "./components/HelloWorld";
import moment from "moment";
moment.locale("id");

export default {
  name: "App",
  components: {
    HelloWorld,
  },
  methods: {
    moment,

    async exportExcel() {
      const workbook = new ExcelJS.Workbook();
      let worksheet = workbook.addWorksheet("ExcelJS sheet", {
        // pageSetup: {
        //   horizontalCentered: true,
        //   columns: 8,
        //   verticalCentered: true,
        //   paperSize: 9,
        //   orientation: "portrait",
        //   margins: {
        //     left: 0.3149606,
        //     right: 0.3149606,
        //     top: 0.3543307,
        //     bottom: 0.3543307,
        //     header: 0.3149606,
        //     footer: 0.3149606,
        //   },
        // },
      });

      const headerRowHasil = worksheet.getRow(1);
      headerRowHasil.height = 30;
      worksheet.mergeCells("A1:AI1");

      headerRowHasil.getCell(1).value =
        "Hasil Monitoring BMKGSoft Bulan " + moment().format(" MMMM  YYYY");
      headerRowHasil.getCell(1).font = {
        name: "Dyuthi",
        size: 22,
        fontWeight: "bold",
        color: "#2acaea",
      };
      headerRowHasil.getCell(1).alignment = {
        horizontal: "center",
        vertical: "middle",
      };

      const headerRowBalai = worksheet.getRow(3);
      headerRowBalai.height = 30;
      worksheet.mergeCells("A3:AI3");

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

      const headerRowTipe = worksheet.getRow(5);
      headerRowTipe.height = 20;
      worksheet.mergeCells("A5:AI5");

      headerRowTipe.getCell(1).value = "Tipe FORM: Me-48";
      headerRowTipe.getCell(1).font = { name: "Segoe UI Light", size: 14 };
      headerRowTipe.getCell(1).alignment = { horizontal: "left" };

      worksheet.getCell("A6:A7").alignment = {
        horizontal: "center",
        vertical: "middle",
      };
      worksheet.mergeCells("A6:A7");

      const customCell3 = worksheet.getCell("A6");
      customCell3.value = "No";

      worksheet.mergeCells("B6:B7");
      worksheet.getCell("B6:B7").alignment = {
        horizontal: "center",
        vertical: "middle",
      };
      const customCell4 = worksheet.getCell("B6");
      customCell4.value = "Stasiun";

      const customCell5 = worksheet.getCell("C6");
      worksheet.mergeCells("C6:Z6");
      worksheet.getCell("C6:Z6").alignment = {
        horizontal: "center",
        vertical: "middle",
      };
      customCell5.value = "Tanggal";

      const customCellTanggal1 = worksheet.getCell("C7");
      customCellTanggal1.value = "1";
      worksheet.getCell("C7").alignment = {
        horizontal: "center",
        vertical: "middle",
      };

      const customCellTanggal2 = worksheet.getCell("D7");
      customCellTanggal2.value = "2";
      worksheet.getCell("D7").alignment = {
        horizontal: "center",
        vertical: "middle",
      };

      const customCellTanggal3 = worksheet.getCell("E7");
      customCellTanggal3.value = "3";
      worksheet.getCell("E7").alignment = {
        horizontal: "center",
        vertical: "middle",
      };

      const customCellTanggal4 = worksheet.getCell("F7");
      customCellTanggal4.value = "4";
      worksheet.getCell("F7").alignment = {
        horizontal: "center",
        vertical: "middle",
      };

      const customCellTanggal5 = worksheet.getCell("G7");
      customCellTanggal5.value = "5";
      worksheet.getCell("G7").alignment = {
        horizontal: "center",
        vertical: "middle",
      };

      const customCellTanggal6 = worksheet.getCell("H7");
      customCellTanggal6.value = "6";
      worksheet.getCell("H7").alignment = {
        horizontal: "center",
        vertical: "middle",
      };

      const customCellTanggal7 = worksheet.getCell("I7");
      customCellTanggal7.value = "7";
      worksheet.getCell("I7").alignment = {
        horizontal: "center",
        vertical: "middle",
      };

      const customCellTanggal8 = worksheet.getCell("J7");
      customCellTanggal8.value = "8";
      worksheet.getCell("J7").alignment = {
        horizontal: "center",
        vertical: "middle",
      };

      const customCellTanggal9 = worksheet.getCell("K7");
      customCellTanggal9.value = "9";
      worksheet.getCell("K7").alignment = {
        horizontal: "center",
        vertical: "middle",
      };

      const customCellTanggal10 = worksheet.getCell("L7");
      customCellTanggal10.value = "10";
      worksheet.getCell("L7").alignment = {
        horizontal: "center",
        vertical: "middle",
      };

      const customCellTanggal11 = worksheet.getCell("M7");
      customCellTanggal11.value = "11";
      worksheet.getCell("M7").alignment = {
        horizontal: "center",
        vertical: "middle",
      };

      const customCellTanggal12 = worksheet.getCell("N7");
      customCellTanggal12.value = "12";
      worksheet.getCell("N7").alignment = {
        horizontal: "center",
        vertical: "middle",
      };

      const customCellTanggal13 = worksheet.getCell("O7");
      customCellTanggal13.value = "13";
      worksheet.getCell("O7").alignment = {
        horizontal: "center",
        vertical: "middle",
      };

      const customCellTanggal14 = worksheet.getCell("P7");
      customCellTanggal14.value = "14";
      worksheet.getCell("P7").alignment = {
        horizontal: "center",
        vertical: "middle",
      };

      const customCellTanggal15 = worksheet.getCell("Q7");
      customCellTanggal15.value = "15";
      worksheet.getCell("Q7").alignment = {
        horizontal: "center",
        vertical: "middle",
      };

      const customCellTanggal16 = worksheet.getCell("R7");
      customCellTanggal16.value = "16";
      worksheet.getCell("R7").alignment = {
        horizontal: "center",
        vertical: "middle",
      };
      const customCellTanggal17 = worksheet.getCell("S7");
      customCellTanggal17.value = "17";
      worksheet.getCell("S7").alignment = {
        horizontal: "center",
        vertical: "middle",
      };

      const customCellTanggal18 = worksheet.getCell("T7");
      customCellTanggal18.value = "18";
      worksheet.getCell("T7").alignment = {
        horizontal: "center",
        vertical: "middle",
      };

      const customCellTanggal19 = worksheet.getCell("U7");
      customCellTanggal19.value = "19";
      worksheet.getCell("U7").alignment = {
        horizontal: "center",
        vertical: "middle",
      };

      const customCellTanggal20 = worksheet.getCell("V7");
      customCellTanggal20.value = "20";
      worksheet.getCell("V7").alignment = {
        horizontal: "center",
        vertical: "middle",
      };

      const customCellTanggal21 = worksheet.getCell("W7");
      customCellTanggal21.value = "21";
      worksheet.getCell("W7").alignment = {
        horizontal: "center",
        vertical: "middle",
      };

      const customCellTanggal22 = worksheet.getCell("X7");
      customCellTanggal22.value = "22";
      worksheet.getCell("X7").alignment = {
        horizontal: "center",
        vertical: "middle",
      };

      const customCellTanggal23 = worksheet.getCell("Y7");
      customCellTanggal23.value = "23";
      worksheet.getCell("Y7").alignment = {
        horizontal: "center",
        vertical: "middle",
      };

      const customCellTanggal24 = worksheet.getCell("Z7");
      customCellTanggal24.value = "24";
      worksheet.getCell("Z7").alignment = {
        horizontal: "center",
        vertical: "middle",
      };

      const customCellTanggal25 = worksheet.getCell("AA7");
      customCellTanggal25.value = "25";
      worksheet.getCell("AA7").alignment = {
        horizontal: "center",
        vertical: "middle",
      };

      const customCellTanggal26 = worksheet.getCell("AB7");
      customCellTanggal26.value = "26";
      worksheet.getCell("AB7").alignment = {
        horizontal: "center",
        vertical: "middle",
      };

      const customCellTanggal27 = worksheet.getCell("AC7");
      customCellTanggal27.value = "27";
      worksheet.getCell("AC7").alignment = {
        horizontal: "center",
        vertical: "middle",
      };

      const customCellTanggal28 = worksheet.getCell("AD7");
      customCellTanggal28.value = "28";
      worksheet.getCell("AD7").alignment = {
        horizontal: "center",
        vertical: "middle",
      };

      const customCellTanggal29 = worksheet.getCell("AE7");
      customCellTanggal29.value = "29";
      worksheet.getCell("AE7").alignment = {
        horizontal: "center",
        vertical: "middle",
      };

      const customCellTanggal30 = worksheet.getCell("AF7");
      customCellTanggal30.value = "30";
      worksheet.getCell("AF7").alignment = {
        horizontal: "center",
        vertical: "middle",
      };

      const customCellTanggal31 = worksheet.getCell("AG7");
      customCellTanggal31.value = "31";
      worksheet.getCell("AG7").alignment = {
        horizontal: "center",
        vertical: "middle",
      };
      worksheet.mergeCells("AH6:AH7");
      worksheet.getCell("AH6:AH7").alignment = {
        horizontal: "center",
        vertical: "middle",
        wrapText: true,
      };
      worksheet.columns = [{ key: "AH", width: 250 }];

      const customCellRataDataMasuk = worksheet.getCell("AH6");
      customCellRataDataMasuk.value = "Rata-rata Data Masuk";
      worksheet.mergeCells("AI6:AI7");
      worksheet.getCell("AI6:AI7").alignment = {
        horizontal: "center",
        vertical: "middle",
        wrapText: true,
      };
      const customCellPenyebabDataTidakMasuk = worksheet.getCell("AI6");
      customCellPenyebabDataTidakMasuk.value = "Penyebab Data Tidak Masuk";
      // worksheet.getRow(9).values = [
      //   "No",
      //   "Stasiun",
      //   "1",
      //   "2",
      //   "3",
      //   "4",
      //   "5",
      //   "6",
      //   "7",
      //   "8",
      //   "9",
      //   "10",
      //   "11",
      //   "12",
      //   "13",
      //   "14",
      //   "15",
      //   "16",
      //   "17",
      //   "18",
      //   "19",
      //   "20",
      //   "21",
      //   "22",
      //   "23",
      //   "24",
      //   "25",
      //   "26",
      //   "27",
      //   "28",
      //   "29",
      //   "30",
      //   "31",
      //   "Rata - Rata Data Masuk",
      //   "Penyebab Data Tidak Masuk",
      // ];

      worksheet.columns = [
        { key: "ID", width: 7, alignment: "center" },
        { key: "Stasiun", width: 30 },
        { key: "1", width: 10 },
        { key: "2", width: 10 },
        { key: "3", width: 10 },
        { key: "4", width: 10 },
        { key: "5", width: 10 },
        { key: "6" },
        { key: "7" },
        { key: "8" },
        { key: "9" },
        { key: "10" },
        { key: "11" },
        { key: "12" },
        { key: "13" },
        { key: "14" },
        { key: "15" },
        { key: "16" },
        { key: "17" },
        { key: "18" },
        { key: "19" },
        { key: "20" },

        { key: "21" },
        { key: "22" },
        { key: "23" },
        { key: "24" },
        { key: "25" },
        { key: "26" },
        { key: "27" },
        { key: "28" },
        { key: "29" },
        { key: "30" },
        { key: "31" },

        { key: "DataMasuk", width: 20 },
        { key: "DataTidakMasuk", width: 30 },
      ];
      worksheet.addRow({
        ID: 1,
        Stasiun: "John Doe",
        1: 101,
        2: 102,
        3: 103,
        4: 104,
        5: 105,
        6: 106,
        7: 107,
        8: 108,
        9: 109,
        10: 110,

        11: 111,
        12: 112,
        13: 113,
        14: 114,
        15: 115,
        16: 116,
        17: 117,
        18: 118,
        19: 119,
        20: 120,

        21: 121,
        22: 122,
        23: 123,
        24: 124,
        25: 125,
        26: 126,
        27: 127,
        28: 128,
        29: 129,
        30: 130,
        31: 131,

        DataMasuk: 200,
        DataTidakMasuk: "Rusakd",
      });

      // FOOTRER
      const rowCount = worksheet.rowCount;
      worksheet.getCell(`A${rowCount + 1}:AH${rowCount + 1}`).alignment = {
        horizontal: "center",
        vertical: "middle",
        wrapText: true,
      };
      worksheet.getCell(`A${rowCount + 1}`).font = { bold: true };
      worksheet.mergeCells(`A${rowCount + 1}:AG${rowCount + 1}`);
      worksheet.getCell(`A${rowCount + 1}:AG${rowCount + 1}`).value = "TOTAL";
      worksheet.getCell(`A${rowCount}:AH${rowCount}`).alignment = {
        horizontal: "center",
        vertical: "middle",
      };

      worksheet.mergeCells(`A${rowCount + 4}:D${rowCount + 4}`);
      worksheet.getCell(`A${rowCount + 4}`).value = "Mengetahui.";
      worksheet.mergeCells(`A${rowCount + 5}:D${rowCount + 5}`);
      worksheet.getCell(`A${rowCount + 5}`).value = "Sub Koordinator";
      worksheet.mergeCells(`A${rowCount + 6}:D${rowCount + 6}`);
      worksheet.getCell(`A${rowCount + 6}`).value =
        "Bidang Manajemen Database MKG,";
      worksheet.mergeCells(`AC${rowCount + 5}:AG${rowCount + 5}`);
      worksheet.getCell(`AC${rowCount + 5}`).value =
        "Jakarta, " + moment().format("DD MMMM  YYYY");

      worksheet.mergeCells(`AC${rowCount + 6}:AG${rowCount + 6}`);
      worksheet.getCell(`AC${rowCount + 6}`).value = "Pembuat Laporan, ";

      // iterate over all current cells in this column including empty cells
      const dobCol = worksheet.getColumn(2);
      dobCol.eachCell({ includeEmpty: true }, function (cell /*rowNumber*/) {
        cell.font = {
          name: "Arial Black",
          color: { argb: "FF00FF00" },
          family: 2,
          size: 14,
          italic: true,
        };
      });

      workbook.xlsx
        .writeBuffer()
        .then((buffer) =>
          saveAs(new Blob([buffer]), `${Date.now()}_feedback.xlsx`)
        )
        .catch((err) => console.log("Error writing excel export", err));
    },
  },
};
</script>

<style>
#app {
  font-family: "Avenir", Helvetica, Arial, sans-serif;
  -webkit-font-smoothing: antialiased;
  -moz-osx-font-smoothing: grayscale;
  text-align: center;
  color: #2c3e50;
  margin-top: 60px;
}
button {
  padding: 10px 20px;
  border: 1px solid #ddd;
  color: #333;
  background-color: green;
  border-radius: 4px;
  font-size: 24px;
  font-weight: bold;
  font-family: "微软雅黑", arail;
  cursor: pointer;
}
</style>

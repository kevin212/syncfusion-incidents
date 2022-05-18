<template>
  <div class="control-section">
    <div id="spreadsheet-default-section">
      <div><h1>@syncfusion/ej2-vue-spreadsheet VER 20.1.56</h1></div>
      <div class="spreadsheet-buttons-container">
        <button
        cssClass="e-control"
        v-on:click="copySpreadsheet">Copy Spreadsheet</button>
        <button
        cssClass="e-control"
        v-on:click="copyDataOnly">Copy Data Only</button>
        <button
        cssClass="e-control"
        v-on:click="copySelection">Copy Selection</button>
      </div>
      <ejs-spreadsheet
      v-if="spreadsheetVisible"
      ref="spreadsheet"
      id="spreadsheet"
      :showRibbon="false"
      :showSheetTabs="false"
      :showFormulaBar="false"
      :enableContextMenu="false"
      :select="onSelected"
      :created="spreadsheetCreated"
      :scrollSettings="scrollSettings"
      :isRowColumnHeadersVisible="false">
        <e-sheets>
          <e-sheet
          :name="worksheetName"
          :frozenRows=1
          :frozenColumns=1
          :showHeaders="false"
          :colCount="maxColCount"
          :rowCount="maxRowCount"
          :height="spreadsheetHeight"
          :isProtected="spreadsheetProtected" :protectSettings="{ selectCells: true }">
            <e-ranges>
              <e-range :dataSource="spreadsheetDataSource" :query="query"></e-range>
            </e-ranges>
            <e-rows>
              <e-row>
                <e-cells>
                  <e-cell  v-for="column in columns" :key="column.field" :value="column.headerText"></e-cell>
                </e-cells>
              </e-row>
            </e-rows>
          </e-sheet>
        </e-sheets>
      </ejs-spreadsheet>
    </div>
  </div>
</template>

<script>
import Vue from "vue";
import { Query } from "@syncfusion/ej2-data";
import * as demoData_A from "./demo-data-a.json";
import { SpreadsheetPlugin } from "@syncfusion/ej2-vue-spreadsheet";

let FONT_SIZE_DEFAULT = '10px';
let FONT_SIZE_DESCRIPTION = FONT_SIZE_DEFAULT;
let FONT_SIZE_DATA = FONT_SIZE_DEFAULT;
let FONT_SIZE_HEADER = FONT_SIZE_DEFAULT;

let DATA_COL_WIDTH = 75;

Vue.use(SpreadsheetPlugin);
export default Vue.extend({
  created() {
    this.spreadsheetDataSource = JSON.parse(JSON.stringify(demoData_A.default));
    this.setDataLength();
    this.setFontSizes();
    this.setCellReferences(this.cellReferences, this.spreadsheetDataSource.length);
    this.query = new Query().select(this.queryModel);
    this.buildColumns(this.spreadsheetDataSource);
    this.maxColCount = this.columns.length;
    this.maxRowCount = this.spreadsheetDataSource.length + 2;
    // show the spreadsheet
    this.spreadsheetVisible = true;
  },
  data: () => {
    return {
      query: [],
      columns: [],
      maxColCount: 25,
      maxRowCount: 100,
      spreadsheetHeight: 750,
      spreadsheetDataSource: [],
      worksheetName: 'LogSheet',
      spreadsheetVisible: false,
      spreadsheetProtected: false,
      scrollSettings: {isFinite: true, enableVirtualization: true},
      queryModel: ['meterDescription','hour01','hour02','hour03','hour04','hour05',
      'hour06','hour07','hour08','hour09','hour10','hour11','hour12','hour13','hour14',
      'hour15','hour16','hour17','hour18','hour19','hour20','hour21','hour22','hour23','hour24'],
      cellReferences: {
        "allCells": "",
        "firsDataCell": "A2",
        "firstColumnLetter": "A",
        "lastColumnLetter": "Z",
        "lastColumnHeader": "Z1",
        "lastColumnLetterDst": "AA",
        "lastColumnHeaderDst": "AA1",
        "singleRowTotalCell" : "Z2",
        "singleRowTotalCellDst" : "AA2",
        "dataColumnHeaderCells": "",
        "firstRowNumber": "",
        "lastRowNumber": "",
        "penultimateRowNumber": "",
        "firstDataCell": "",
        "lastDataCell": "",
        "dataCells": "",
        "firstCell": "",
        "lastCell": "",
        "firstColumn": "",
        "lastColumn": "",
        "firstRow": "",
        "lastRow": "",
        "meterDescriptionCells": "",
        "totalsRowLabelCell": "",
        "totalsRowCells": ""
      }
    }
  },
  methods: {
    copySelection() {
      const spreadsheet = this.$refs.spreadsheet;
      spreadsheet.copy(this.lastRangeSelection);
    },
    copySpreadsheet() {
      const spreadsheet = this.$refs.spreadsheet;
      spreadsheet.copy(this.cellReferences.allCells);
    },
    copyDataOnly() {
      const spreadsheet = this.$refs.spreadsheet;
      spreadsheet.copy(this.cellReferences.dataCells);
    },
    onSelected(args) {
      this.lastSelection = args.range;
    },
    spreadsheetCreated() {
      this.addTotalsColumn();
      this.addTotalsRowFormulas(this.buildTotalsRowFormulaConfigs());
      this.styleSpreadsheet();
    },
    objectHasProperty(object, property) {
        return Object.prototype.hasOwnProperty.call(object, property);
    },
    buildQueryModel(object, property) {
        return Object.prototype.hasOwnProperty.call(object, property);
    },
    letterFromNumber(num) {
      let letter = String.fromCharCode(97 + num);
      return letter.toUpperCase();
    },
    queryToObject (){
        var i = 0, retObj = {},
          pair = null, sPageURL = window.location.search.substring(1),
          qArr = sPageURL.split('&');

        for (; i < qArr.length; i++){pair = qArr[i].split('='); retObj[pair[0]] = pair[1]; }
        return retObj;
    },
    setDataLength() {
        const queryObject = this.queryToObject();

        if (queryObject && queryObject.rows) {
          const numberOfRows = parseInt(queryObject.rows);

          if (numberOfRows <= this.spreadsheetDataSource.length) {
            this.spreadsheetDataSource.length = numberOfRows;
          }
        } else {
          this.spreadsheetDataSource.length = 15;
        }
    },
    setFontSizes() {
        const queryObject = this.queryToObject();

        if (queryObject && queryObject.fontSize) {
          FONT_SIZE_DEFAULT = `${queryObject.fontSize}px`;

          FONT_SIZE_DESCRIPTION = FONT_SIZE_DEFAULT;
          FONT_SIZE_DATA = FONT_SIZE_DEFAULT;
          FONT_SIZE_HEADER = FONT_SIZE_DEFAULT;
        }

        if (queryObject && queryObject.fontSizeData) {
          FONT_SIZE_DATA = `${queryObject.fontSizeData}px`;
        }

        if (queryObject && queryObject.fontSizeHeader) {
          FONT_SIZE_HEADER = `${queryObject.fontSizeHeader}px`;
        }

        if (queryObject && queryObject.fontSizeDesc) {
          FONT_SIZE_DESCRIPTION = `${queryObject.fontSizeDesc}px`;
        }

        if (queryObject && queryObject.dataColWidth) {
          DATA_COL_WIDTH = `${queryObject.dataColWidth}`;
        }

        console.warn(`setFontSizes - ${FONT_SIZE_DEFAULT}`);
        console.dir(queryObject);
    },
    addTotalsColumn() {
      const spreadsheet = this.$refs.spreadsheet;

      this.spreadsheetDataSource.forEach((rowObject, index) => {
          const row = index + 2;
          const formula = `=SUM(B${row}:Y${row})`;
          const config = {isLocked: true, formula};
          spreadsheet.updateCell(config, `Z${row}`);
      });
    },
    addTotalsRowFormulas(configs) {
      const spreadsheet = this.$refs.spreadsheet;
      spreadsheet.insertRow(this.spreadsheetDataSource.length + 1);

      configs.forEach(config => {
        spreadsheet.updateCell(config.props, config.cell);
      });
    },
    alphabetPosition(text) {
      let result = '';

      for (let i = 0; i < text.length; i++) {
        let code = text.toUpperCase().charCodeAt(i);
        if (code > 64 && code < 91) {
          result += (code - 64) + ' ';
        }
      }

      return result.slice(0, result.length - 1);
    },
    buildTotalsRowFormulaConfigs() {
      const configs = [];
      const firstLetter = 'B';
      const lastLetter = 'Z';
      const firstLetterIndex = this.alphabetPosition(firstLetter);
      const lastLetterIndex = this.alphabetPosition(lastLetter);

      for (let i = firstLetterIndex; i <= lastLetterIndex; i++ ) {
        const letter = this.letterFromNumber(i - 1);

        configs.push({
          cell: `${letter}${this.spreadsheetDataSource.length + 2}`,
          props : { formula: `=SUM(${letter}2:${letter}${this.spreadsheetDataSource.length + 1})`, isLocked: false}
        });
      }

      return configs;
    },
    buildColumns(data) {
      const hourText = 'hour';
      const firstElement = data[0];

      // build meter description column
      this.columns.push({
        width:  50,
        field: 'name',
        textAlign: 'Left',
        headerText: 'Meter Description'
      });

      // build 'hourXX' columns
      for (let prop in firstElement) {
        // target only 'hourXX' props
        if (prop.indexOf(hourText) > -1) {
          const headerTextArr = prop.split(hourText);

          this.columns.push({
            width:  150,
            field: prop,
            textAlign: 'center',
            headerText: headerTextArr[1]
          });
        }
      }

      // build row total column
      this.columns.push({
        width: 150,
        field: 'rowTotal',
        headerText: 'Total',
        textAlign: 'center'
      });
    },
    styleSpreadsheet() {
      const spreadsheet = this.$refs.spreadsheet;
      const lastRowIndex = this.spreadsheetDataSource.length + 2;
      const headersConfig = {fontWeight: 'bold', fontSize: FONT_SIZE_HEADER, backgroundColor: '#eef3f5'};
      const totalsConfig = {fontWeight: 'bold', backgroundColor: '#fefee1'};

      spreadsheet.cellFormat(headersConfig, 'A1:Z1');
      spreadsheet.cellFormat(headersConfig, `A2:A${this.spreadsheetDataSource.length + 1}`);

      spreadsheet.cellFormat({textAlign: 'left', fontSize: FONT_SIZE_DESCRIPTION, fontFamily: 'Arial'}, `A1:A${lastRowIndex}`);
      spreadsheet.cellFormat({textAlign: 'center', fontFamily: 'Arial'}, `A1:Z${lastRowIndex}`);


      spreadsheet.cellFormat(totalsConfig, `Z2:Z${lastRowIndex}`);
      spreadsheet.cellFormat(totalsConfig, `B${lastRowIndex}:Z${this.spreadsheetDataSource.length + 2}`);
      spreadsheet.numberFormat('0.###', `B2:Z${this.spreadsheetDataSource.length + 2}`);
      spreadsheet.cellFormat({backgroundColor: '#f5f0ee', textAlign: 'left'}, `A2:A${this.spreadsheetDataSource.length + 1}`);

      spreadsheet.cellFormat({fontSize: FONT_SIZE_DATA}, `B2:Z${this.spreadsheetDataSource.length + 2}`);

      // set the meter description width
      spreadsheet.setColWidth(150, 0, 0);

      // set data colum widths
      for (let i = 1; i < 25; i++) {
        spreadsheet.setColWidth(DATA_COL_WIDTH, i, 0);
      }
    },
    setCellReferences(cellReferences, dataSourceLength) {
      const cells = cellReferences;
      const firstRowNumber = 1;
      const lastRowNumber = dataSourceLength + 2;
      const penultimateColumLetterIndex = this.alphabetPosition(cells.lastColumnLetter) - 2;
      const penultimateColumLetter = this.letterFromNumber(penultimateColumLetterIndex);

      cells.firstRowNumber = firstRowNumber;
      cells.lastRowNumber = lastRowNumber;
      cells.penultimateRowNumber = lastRowNumber - 1;

      cells.firstCell = `${cells.firstColumnLetter}${firstRowNumber}`;
      cells.lastCell = `${cells.lastColumnLetter}${lastRowNumber}`;
      cells.allCells = `${cells.firstCell}:${cells.lastCell}`;

      cells.firstColumn = `${cells.firstCell}:${cells.firstColumnLetter}${lastRowNumber}`;
      cells.lastColumn = `${cells.lastColumnLetter}${firstRowNumber}:${cells.lastCell}`;

      cells.firstRow = `${cells.firstColumnLetter}${firstRowNumber}:${cells.lastColumnLetter}${firstRowNumber}`;
      cells.lastRow = `${cells.firstColumnLetter}${lastRowNumber}:${cells.lastCell}`;

      cells.firstDataCell = `B2`
      cells.lastDataCell = `${penultimateColumLetter}${dataSourceLength + 1}`;
      cells.dataCells = `${cells.firstDataCell}:${cells.lastDataCell}`;

      cells.meterDescriptionCells = `A2:A${dataSourceLength - 1}`;
      cells.totalsRowLabelCell = `A${dataSourceLength}`;

      cells.totalsRowCells = `${cells.lastColumnLetter}2:${cells.lastColumnLetter}${dataSourceLength - 1}`;

      cells.dataColumnHeaderCells = `${cells.firstColumnLetter}${firstRowNumber}:${cells.lastColumnLetter}${firstRowNumber}`;
    }
  }
});
</script>

<style>
#spreadsheet-default-section .spreadsheet-buttons-container{
  margin-bottom: 10px;
}

#spreadsheet-default-section .spreadsheet-buttons-container button{
  margin-right: 10px;
}
</style>


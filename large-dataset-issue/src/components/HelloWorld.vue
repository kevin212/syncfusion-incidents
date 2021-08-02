<template>
  <div class="control-section">
    <div id="spreadsheet-default-section">
      <ejs-spreadsheet
      v-if="spreadsheetVisible"
      ref="spreadsheet"
      id="spreadsheet"
      :showFormulaBar="false"
      :showRibbon="false"
      :showSheetTabs="false"
      :enableContextMenu="false"
      :isRowColumnHeadersVisible="false"
      :scrollSettings="scrollSettings"
      :created="spreadsheetCreated">
        <e-sheets>
          <e-sheet
          :name="worksheetName"
          :frozenRows=1
          :frozenColumns=1
          :colCount="maxColCount"
          :rowCount="maxRowCount"
          :height="spreadsheetHeight"
          :showHeaders="false"
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
import * as demoData from "./large-dataset.json";
import { SpreadsheetPlugin } from "@syncfusion/ej2-vue-spreadsheet";

const COL_WIDTH_DATA_CELL = 100;
const COL_WIDTH_ROW_TOTAL = 200;
const COL_WIDTH_METER_DESCRIPTION = 250;

Vue.use(SpreadsheetPlugin);
export default Vue.extend({
  created() {
    this.spreadsheetDataSource = JSON.parse(JSON.stringify(demoData.default));
    this.processQueryString();
    this.query = new Query().select(this.queryModel);
    this.buildColumns(this.spreadsheetDataSource);
    this.maxColCount = this.columns.length;
    this.maxRowCount = this.spreadsheetDataSource.length + 2;
    setTimeout(() => {this.spreadsheetVisible = true;}, 1000);
  },
  data: () => {
    return {
      worksheetName: 'LogSheet',
      spreadsheetVisible: false,
      spreadsheetHeight: 500,
      query: [],
      maxColCountx: 25,
      maxRowCount: 100,
      spreadsheetProtected: false,
      spreadsheetDataSource: [],
      rowIndex: 30,
      colIndex: 4,
      columns: [],
      scrollSettings: {isFinite: true, enableVirtualization: true},
      queryModel: [
        'meterDescription',
        'hour01',
        'hour02',
        'hour03',
        'hour04',
        'hour05',
        'hour06',
        'hour07',
        'hour08',
        'hour09',
        'hour10',
        'hour11',
        'hour12',
        'hour13',
        'hour14',
        'hour15',
        'hour16',
        'hour17',
        'hour18',
        'hour19',
        'hour20',
        'hour21',
        'hour22',
        'hour23',
        'hour24'
      ]
    }
  },
  methods: {
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
    styleSpreadsheet() {
      const spreadsheet = this.$refs.spreadsheet;
      const lastRowIndex = this.spreadsheetDataSource.length + 2;
      const headersConfig = {fontWeight: 'bold', fontSize: '12px', backgroundColor: '#eef3f5'};
      const totalsConfig = {fontWeight: 'bold', backgroundColor: '#fefee1'};

      spreadsheet.cellFormat(headersConfig, 'A1:Z1');
      spreadsheet.cellFormat(headersConfig, `A2:A${this.spreadsheetDataSource.length + 1}`);
      spreadsheet.cellFormat({textAlign: 'left', fontSize: '10px', fontFamily: 'Arial'}, `A1:Z${lastRowIndex}`);
      spreadsheet.cellFormat(totalsConfig, `Z2:Z${lastRowIndex}`);
      spreadsheet.cellFormat(totalsConfig, `B${lastRowIndex}:Z${this.spreadsheetDataSource.length + 2}`);
      spreadsheet.numberFormat('###,###.###', `B2:Z${this.spreadsheetDataSource.length + 1}`);
    },
    letterFromNumber(num) {
      let letter = String.fromCharCode(97 + num)
      return letter.toUpperCase();
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
    processQueryString() {
      const queryStringObject = this.queryToObject();
      const dataSource = this.spreadsheetDataSource;

      if (queryStringObject && queryStringObject.dataRows){
        console.warn('queryStringObject ->');
        console.dir(queryStringObject);
        const totaldDtaRows = parseInt(queryStringObject.dataRows, 10);

        if (totaldDtaRows && totaldDtaRows < dataSource.length){
          dataSource.length = totaldDtaRows;
          console.warn(`Data size: ${dataSource.length} elements`);	
        } else {
          console.warn(`queryString.dataRows equals or exceeds actual data size. Using actual data size. (${dataSource.length} elements)`);
        }
      }
    },
    queryToObject (){
        var
        i = 0,
        retObj = {},
        pair = null,
        sPageURL = window.location.search.substring(1),
        qArr = sPageURL.split('&');

        for (; i < qArr.length; i++){
            pair = qArr[i].split('=');
            retObj[pair[0]] = pair[1];
        }

        return retObj;
    },
    buildColumns(data) {
      const hourText = 'hour';
      const firstElement = data[0];

      this.columns.push({
        field: 'name',
        textAlign: 'Left',
        headerText: 'Meter Description',
        width:  COL_WIDTH_METER_DESCRIPTION
      });

      for (let prop in firstElement) {
        // target only "hourXX" props
        if (prop.indexOf(hourText) > -1) {
          const headerTextArr = prop.split(hourText);

          this.columns.push({
            field: prop,
            textAlign: 'center',
            headerText: headerTextArr[1],
            width:  COL_WIDTH_DATA_CELL
          });
        }
      }

      this.columns.push({
        field: 'rowTotal',
        headerText: 'Total',
        textAlign: 'center',
        width: COL_WIDTH_ROW_TOTAL
      });
    }
  }
});
</script>


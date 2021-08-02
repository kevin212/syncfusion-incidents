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
    this.spreadsheetDataSource = demoData.default;
    this.query = new Query().select(this.queryModel);
    this.buildColumns(this.spreadsheetDataSource);
    this.maxColCount = this.columns.length;
    this.maxRowCount = this.spreadsheetDataSource.length + 1;
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
      spreadsheetProtected: true,
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
      this.styleSpreadsheet();
    },
    objectHasProperty(object, property) {
        return Object.prototype.hasOwnProperty.call(object, property);
    },
    styleSpreadsheet() {
      const spreadsheet = this.$refs.spreadsheet;

      spreadsheet.cellFormat({fontWeight: 'bold', fontSize: '12px'}, 'A1:Z1');
      spreadsheet.cellFormat({fontWeight: 'bold', fontSize: '12px'}, `A2:A${this.spreadsheetDataSource.length + 1}`);
      spreadsheet.cellFormat({textAlign: 'left', fontSize: '10px', fontFamily: 'Arial'}, `A1:Z${this.spreadsheetDataSource.length + 1}`);
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


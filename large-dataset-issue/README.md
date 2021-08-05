# syncfusion-incidents
Demo Code for Syncfusion Large Dataset Issue

## Local Setup

```
npm install
npm run serve
```
## Local Testing Scenarios:

http://localhost:8080?sort=asc - Sort data acending

http://localhost:8080?sort=desc - Sort data decending

http://localhost:8080?dataRows=10 - Load 10 data rows

http://localhost:8080?dataRows=150 - Load 150 data rows

http://localhost:8080?sort=asc&dataRows=225 - Load 225 data rows, Sort data acending

ETC...

### Alternate Data Sets:

http://localhost:8080?dataSet=b

http://localhost:8080?dataSet=c

http://localhost:8080?dataSet=b&dataRows=10

http://localhost:8080?dataSet=c&dataRows=10

http://localhost:8080?dataSet=b&dataRows=50

http://localhost:8080?dataSet=c&dataRows=50

http://localhost:8080?dataSet=b&dataRows=200

http://localhost:8080?dataSet=c&dataRows=200

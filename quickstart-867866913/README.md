# quickstart

> A Vue.js project

## Build Setup

**install dependencies** - yarn install

**serve with hot reload at localhost:8080** - yarn dev

**You can set the size of the data**:

http://localhost:8080?rows=50 - 50 rows
http://localhost:8080?rows=100 - 100 rows
http://localhost:8080?rows=500 - 500 rows
http://localhost:8080?rows=1000 - 1000 rows (mnax)
...etc

-----------------------------------------------------
# VIDEO DEMONSTRATION OF THE SETUP (following your instructions)

[https://kevinchisholm.com/test/gmp/quickstart-867866913/custom-pkg-setup-12-13-2021.mp4](https://kevinchisholm.com/test/gmp/quickstart-867866913/custom-pkg-setup-12-13-2021.mp4)

# Issues

## Issue # 1: - When there are more than 50 rows, and I select the ENTIRE spreadsheet manually, and then COPY using "C TRL + C" (or "⌘ + C" for Mac), pasting into MS Excel does not work at all AND the page freezes.

VIDEO DEMONSTRATION

[https://kevinchisholm.com/test/gmp/quickstart-867866913/copy-spreadsheet-manually-12-13-2021.mp4](https://kevinchisholm.com/test/gmp/quickstart-867866913/copy-spreadsheet-manually-12-13-2021.mp4)

## Issue # 2 - When I COPY the ENTIRE spreadsheet programmatically, and there are more than 50 rows, pasting into MS Excel does not work properly.

VIDEO DEMONSTRATION: 

[https://kevinchisholm.com/test/gmp/quickstart-867866913/copy-spreadsheet-programmatically-12-13-2021.mp4](https://kevinchisholm.com/test/gmp/quickstart-867866913/copy-spreadsheet-programmatically-12-13-2021.mp4)

## 12/14/2021

It seems that if there are more rows than the UI can show, then programmatic copy & paste does not work correctly, unless you scroll **ALL** the way to the bottom of the spreadsheet so that all rows have rendered in the UI. 

**More specifically, it seems to pertain to cells that have been added programmatically, but not rendered in the UI yet**. 

THIS VIDEO DEMONSTRATION MAY MAKE IT EASIER TO EXPLAIN:

[https://kevinchisholm.com/test/gmp/quickstart-867866913/copy-spreadsheet-programmatically-12-14-2021.mp4](https://kevinchisholm.com/test/gmp/quickstart-867866913/copy-spreadsheet-programmatically-12-14-2021.mp4)


**NOTE**: Manual copy/paste still does not work with a large data-set and it still locks-up the browser

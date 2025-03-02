# ExcelOutPut
Focus primarily on OOP structure design. Write the Program class in pseudo code.


## 功能
透過 excelExporter class 匯出 excel 檔案。
假設return XSSWorkBook 就算完成匯出。

## 前提
假設一個Program就是一個頁面。
`DataSet`:
UI畫面有顯示出來的部份資料，其他資料需要透過 `dbMethod` 從資料庫取得。


## 設計理念
有可能很多頁面會需要匯出excel的功能，所以設計了一個`ExcelExporter` class
讓其他頁面可以使用。

其他頁面要用的這個功能只需要使用 `AddColumn`和實作 `IRetriever`就好。 

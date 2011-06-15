Running Instruction:

1. 將刷卡紀錄匯出成xls檔, 放到桌面report目錄下
2. 編輯report目錄下的conf.ini
   Ex:
   2011 6	// 欲產生報表的年 月, 以空格隔開
   9 40		// 定義遲到的時 分, 以空格隔開
   22 23 24	// 實習生的工號, 以空格隔開
3. 雙點start_report.bat執行script
4. 報告產生於同目錄下201X_X_Report.csv

Running Environment:

1. Download Ruby windows installer - 1.9.2-p180
2. $gem install roo builder google-spreadsheet-ruby zip spreadsheet
3. Set windows path:
   $PATH=$PATH;C:\Ruby192\bin
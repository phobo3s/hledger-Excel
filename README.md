# hledger-Excel
hledger journal parser with Excel and VBA

Hledger and Excel Together

Dear PTA enthusiasts,

After extensive efforts, I have created an Excel-VBA project that processes data from my bank and generates a .txt file in Hledger format from the Excel tables I prepared. Since I learned VBA on my own, it’s not professional, so please bear with me.

The reasons why I used Excel, which may not align well with the philosophy of PTA, are as follows:

Writing in Excel tables seemed easier than writing in a text file.

I thought I could convert the file to CSV format and share it with other applications (for example, Portfolio Performance "https://www.portfolio-performance.info/"). This way, I could more easily create charts and perform other analyses.

I wanted to track the quantity and unit prices of the stocks and funds in my portfolio. The buy and sell transactions that I wrote appropriately in Excel cells are converted to Hledger format. I can track stocks and funds like inventory.

I can calculate my net profit using FIFO for sales transactions. I handle all buy and sell transactions within a "Dictionary" object and deplete them like inventory.

For stock splits, I needed to update the historical quantity and price information backwards so that the real-time prices from the internet matched my current portfolio.

I used Excel for these reasons. In the distant future, I might transfer this work to Google Sheets so it can work on the web as well, but this will be VERY DIFFICULT.

Please excuse my poor English, it is not my native language.

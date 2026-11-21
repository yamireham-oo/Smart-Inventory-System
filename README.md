# Smart-Inventory-System
An Excel and Power BI project that tracks office supply stock, records stock movements, and shows which items need to be reordered.

<img width="978" height="547" alt="Screenshot 2026-11-20 122153" src="https://github.com/user-attachments/assets/3e2ec3d9-bdef-43b1-908c-dcc873b1fb34" />
<img width="985" height="547" alt="Screenshot 2026-11-20 122308" src="https://github.com/user-attachments/assets/f7cf3acc-a5c5-4792-8d4b-2810a04b4641" />

[Download the workbook](Smart-Inventory-System.xlsx)

What it does

- Records stock movements (IN / OUT) in a transactions table, with a dropdown to keep entries consistent.
- Calculates total in, total out and current stock for each product automatically.
- Compares current stock with target stock and labels each product Good, Low Stock or Out Of Stock.
- Calculates how many units to reorder.
- Shows the results in a Power BI dashboard with stock status cards, a status chart, a current vs target chart and a reorder table, filtered by a stock status slicer.

Formulas

Total In and Total Out add up the quantity for each product by movement type.

```excel
=SUMIFS(Table1[Quantity],Table1[Product ID],[@[Product ID]],Table1[Movement Type],"IN")
=SUMIFS(Table1[Quantity],Table1[Movement Type],"OUT",Table1[Product ID],[@[Product ID]])
```

Current Stock starts from the opening stock.

```excel
=[@[Opening Stock]]+[@[Total In]]-[@[Total Out]]
```

Stock Status compares current stock with the target.

```excel
=IFS([@[Current Stock]]<=0,"Out Of Stock",[@[Current Stock]]<[@[Target Stock]],"Low Stock",[@[Current Stock]]>=[@[Target Stock]],"Good")
```

Reorder Quantity shows how much to order to reach the target, and never goes below zero.

```excel
=MAX(0,[@[Target Stock]]-[@[Current Stock]])
```

Dashboard insights

The inventory has 12 products: 2 Good, 9 Low Stock and 1 Out Of Stock. Most items are below target, and the reorder table shows how many units to order for each product.

Skills: Excel Tables, SUMIFS, IFS, MAX, data validation, Power BI, dashboard design

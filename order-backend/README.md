# PattEcks! pre-orders → Google Sheets

Each pre-order from `order.html` becomes one row in a Google Sheet. The payment slip is saved to a Google Drive folder called "PattEcks slips", and the row links to it.

## Set up (once, about 10 minutes)

1. Create a new Google Sheet (for example "PattEcks Orders").
2. In the sheet: **Extensions → Apps Script**.
3. Delete what is in `Code.gs`, paste the contents of `Code.gs` from this folder, and save.
4. **Deploy → New deployment**. Type: **Web app**.
   - Execute as: **Me**
   - Who has access: **Anyone**
5. Press **Deploy**, allow the permissions Google asks for (Sheets and Drive), and copy the **Web app URL** (it ends in `/exec`).
6. Put that URL in the shop settings of the order page (`orderUrl`), or send it to Claude to do it.

The "Orders" tab is created by the first order. Use the **Status** column to track each order (New → Paid → Shipped).

If you change `Code.gs` later, deploy again with **Manage deployments → Edit → New version**, so the URL stays the same.

## Privacy

The sheet and the slip folder hold customers' names, phone numbers and addresses. Keep both private to you, and delete old rows when orders are done and no longer needed.

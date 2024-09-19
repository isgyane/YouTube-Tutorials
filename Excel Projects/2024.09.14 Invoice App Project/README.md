# Excel Invoice Application

[Watch Full Tutorial Here](https://youtu.be/MWcTV3-vDMg)  
[Download Starter File](https://raw.githubusercontent.com/isgyane/YouTube-Tutorials/main/Excel%20Projects/2024.09.14%20Invoice%20App%20Project/Invoice%20Starter.xlsx)


This tutorial guides you through the process of creating a fully functional and automated invoice application in Microsoft Excel. It covers key steps such as designing the invoice layout, applying necessary formatting, and using essential formulas to calculate totals, discounts, and more. The tutorial also demonstrates how to apply conditional formatting, manage print settings for a professional look, and ensure the invoice is ready for export as a PDF. With this step-by-step guide, you can streamline the invoicing process and easily generate professional invoices for your business.

## Table of Contents

1. [Setting Up the Invoice](#setting-up-the-invoice)
   - 2.1 [Designing the Layout](#designing-the-layout)
   - 2.2 [Formatting the Cells](#formatting-the-cells)
2. [Formulas Used](#formulas-used)
   - 3.1 [Total Calculation](#total-calculation)
   - 3.2 [Discount Application](#discount-application)
   - 3.3 [Conditional Formatting for Discounts](#conditional-formatting-for-discounts)
3. [Special Notes and Instructions](#special-notes-and-instructions)
4. [Handling Page Breaks and Printing](#handling-page-breaks-and-printing)
5. [Conclusion](#conclusion)

---

## Setting Up the Invoice

### Designing the Layout

In this chapter, we set up the structure of the invoice, including headers for product details, quantity, price, discount, and total.

1. **Headers:** Create headers like "Product Name," "Quantity," "Unit Price," "Discount," and "Total Price."
2. **Borders:** Ensure the table looks professional by adding borders to all cells that will hold data.

### Formatting the Cells

Use **Merge Cells** and formatting techniques to create special sections for:
- **Invoice Title**
- **Date**
- **Invoice Number**
- **Customer Information**
- **Special Notes** (at the bottom of the invoice for additional instructions).

## Formulas Used

### Total Calculation

To calculate the total cost of each product:
- Use a formula that multiplies the quantity by the unit price.

### Discount Application

To apply a discount on the total price:
- Use an IF statement that checks if a discount is applicable and adjusts the total accordingly.

### Conditional Formatting for Discounts

To ensure that no discount information appears if there is no discount:
1. Select the Discount cell.
2. Go to **Conditional Formatting** > **New Rule**.
3. Set a rule that applies formatting based on the discount value.

## Special Notes and Instructions

At the bottom of the invoice, there is a **Special Notes and Instructions** section that allows users to input any additional information relevant to the customer.

Steps:
1. Merge the required cells across the bottom of the invoice.
2. Add a title such as "Special Notes and Instructions."
3. Apply formatting, such as borders and background color, to distinguish this section.

## Handling Page Breaks and Printing

To ensure the invoice prints on a single page:

1. **Page Break Preview:** Go to **View** > **Page Break Preview** to adjust where Excel will break the page.
2. **Adjust Columns:** Drag the boundaries to fit everything on one page if necessary.
3. **Print Area:** Select the cells to include in the printed version and set the **Print Area**.

Finally, you can export the invoice as a **PDF**:
- Press **Ctrl+P** to access the print menu.
- Choose **Microsoft Print to PDF** and save the file.
- Alternatively, select your printer if you prefer to print a hard copy of the invoice.

## Conclusion

This project demonstrates how to create a structured and automated invoice application using Excel. The flexibility of Excel formulas, cell formatting, and print controls allows users to customize this application according to their specific needs.

For further customization or requests, feel free to [open an issue](https://github.com/your-repo/issues) or contact me.

# Data issues

## store_north.xlsx
- **1. Exact duplicate rows.** 5 orders appear twice with identical values (N-1049, N-1068, N-1118, N-1127, N-1151), so the file has 185 rows for 180 orders.
- **2. Missing customer details.** 4 rows have no customer name and 5 rows have no email.

## store_south.xlsx
- **3. Prices stored as text in three formats.** Values look like `$39.00`, `39.00` and `39.00 USD` (with thousands separators on large amounts, such as `$1,249.00`). 7 rows have no price at all.
- **4. Bad quantity values.** 9 quantities are text with a unit, such as `2 pcs`, and 6 are impossible values (`0` or `-1`).
- **5. Empty rows in the middle of the data.** 2 completely blank rows are mixed in among the orders.

## online_shop.xlsx
- **6. Report title above the header.** The first three rows hold a title, a "Generated" date and a blank line, so the real header is on Excel row 4. Read as-is, every column comes out as `Unnamed: N`.
- **7. Summary rows at the bottom.** A blank row and a `TOTAL` row (85,573.25) follow the last order. They are not orders and would inflate any sum.
- **8. Orders already exported by store_north.** 12 orders with `N-` IDs also appear in the store_north file, so combining the files would count them twice. These copies are **not identical rows**: the date format, name spelling and city spelling differ, so an exact-duplicate check will not find them. They can only be matched by Order ID, after the other columns are cleaned. 7 orders here also have no city (`Location`).

## Across all files
- **9. Different column names for the same data.**

  | Field      | store_north  | store_south     | online_shop     |
  |------------|--------------|-----------------|-----------------|
  | Order ID   | `Order ID`   | `order_id`      | `ID`            |
  | Date       | `Order Date` | `date`          | `Date`          |
  | Customer   | `Customer`   | `customer_name` | `Name`          |
  | Email      | `Email`      | `e-mail`        | `Email Address` |
  | Phone      | `Phone`      | `phone_number`  | `Phone`         |
  | City       | `City`       | `city`          | `Location`      |
  | Product    | `Product`    | `product`       | `Item`          |
  | Quantity   | `Qty`        | `quantity`      | `Qty`           |
  | Unit price | `Unit Price` | `price`         | `Price (USD)`   |

- **10. Three different date formats.** store_north has real Excel dates, store_south has text in day.month.year (`24.03.2026`), and online_shop has text in month/day/year (`2/5/2026`). A value like `2/5` is ambiguous unless you know which file it came from.
- **11. Inconsistent text.** Names, products and cities mix UPPER, lower and Title case, and some have leading, trailing or double spaces. The 7 products appear in 35 different spellings, and the 5 cities in 20 (`NYC`, `New York City`, `NY`, `LA`, `L.A.`, `SF`, `San Fran`, `Chicago, IL`, `Houston TX`, …). 95 of 535 customer names have extra spaces.
- **12. Unreliable contact details.** Phone numbers come in 6 formats (`(555) 123-4567`, `555-123-4567`, `555.123.4567`, `+1 555 123 4567`, `5551234567`, `1-555-123-4567`). 41 emails are invalid because they lack the `@` or the `.com` ending (19 in north, 14 in south, 8 online), and others are in uppercase or padded with spaces.
- **13. Non-data rows break column types.** Because of the blank and `TOTAL` rows, pandas reads online_shop's `Qty` as decimals (`2.0` instead of `2`), and store_south's `quantity` is a mix of numbers and text. Rows that are not orders must be removed first, and only then can columns be converted to the right types.

## Assumptions
- `NY` means New York City, not New York State, and `LA` means Los Angeles, not Louisiana. All other values in the city column are cities, so this is the most likely meaning. With a real client, this would be confirmed with them first.
- Each Order ID identifies one order across all files, so two rows with the same Order ID are the same order.
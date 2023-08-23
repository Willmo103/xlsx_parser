import sys
from app import extract_flagged_sales_orders

# # call the parsing function with the input and output paths from the command line
# columns = [
#     "Warehouse Name",
#     "Sales Order",
#     "SO Line Nbr",
#     "Customer PO ID",
#     "Item ID",
#     "LineDescription",
#     "Book Date",
#     "Promise Date",
#     "Required Date",
#     "Ship Window Close",
#     "weight per item",
#     "Required Qty",
#     "Weight total",
#     "pallet count",
#     "Current Backlog",
#     "CTP",
#     "Outage",
#     "Purchased/ Manufactured",
#     "Action/Comments",
#     "Days From Ship Window End",
#     "Bobby's Date",
# ]

if __name__ == "__main__":
    # for i, col in enumerate(
    #     columns
    # ):  # enumerate allows us to loop over a list and keep track of the index
    #     print(f"{i}: {col}")
    # group_by = int(input("Enter the number of the column to group by: "))
    extract_flagged_sales_orders(sys.argv[1], "output.xlsx")
    # extract_flagged_sales_orders("input.xlsx", "output.xlsx")

#!/usr/bin/env bash
# Test script to demonstrate formula preservation feature

echo "Waiting for app to start..."
sleep 5

BASE_URL="http://localhost:8080/api/excel"

# Check if app is running
if ! curl -s "$BASE_URL/health" > /dev/null 2>&1; then
    echo "❌ App is not running on port 8080"
    exit 1
fi

echo "✅ App is running"
echo ""
echo "Testing formula preservation feature..."
echo "======================================="
echo ""
echo "Sending test data with 'discountedPrice' field:"
echo "(This will be mapped to cell C2 which contains formula =B2*1.1)"
echo ""

read -r -d '' PAYLOAD <<'JSON'
{
  "discountedPrice": 999.99
}
JSON

echo "Request payload:"
echo "$PAYLOAD"
echo ""

response=$(curl -s -w "\nHTTP_STATUS:%{http_code}" -X POST "$BASE_URL/merge" \
    -H "Content-Type: application/json" \
    -d "$PAYLOAD")

status=$(echo "$response" | sed -n 's/.*HTTP_STATUS://p')
echo "Response status: $status"

if [[ "$status" == "200" ]]; then
    output=$(echo "$response" | grep -o '"outputPath"[^"]*"[^"]*"')
    echo "✅ Merge successful"
    echo "Output file: $output"
    echo ""
    echo "Checking formula preservation in output..."
    
    python3 << 'PYTHON_EOF'
import glob
import os
from openpyxl import load_workbook

files = glob.glob('./output/merged_*.xlsx')
if files:
    latest_file = max(files, key=os.path.getctime)
    print(f"Reading: {latest_file}")
    
    wb = load_workbook(latest_file)
    ws = wb.active
    
    # Check C2
    c2_cell = ws['C2']
    print(f"\nCell C2 (should preserve formula):")
    print(f"  Cell type: {c2_cell.data_type}")
    print(f"  Cell value: {c2_cell.value}")
    if c2_cell.data_type == 'f':
        print(f"  ✅ Formula preserved: {c2_cell.value}")
    else:
        print(f"  ⚠️  Formula was overwritten with value")
    
    wb.close()
else:
    print("No merged files found")
PYTHON_EOF
else
    echo "❌ Merge failed"
    echo "$response"
fi

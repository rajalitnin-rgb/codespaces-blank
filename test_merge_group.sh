#!/usr/bin/env bash
# simple script to POST sample JSON to merge service and display result

BASE_URL="http://localhost:8080/api/excel"

# sample payload matching application.yml mappings
read -r -d '' PAYLOAD <<'JSON'
{
  "firstName": "Alice",
  "months": ["Jan", "Feb", "Mar", "Apr", "May", "Jun"],
  "quarters": ["Q1", "Q2", "Q3", "Q4"],
  "salesMatrix": [
    [100, 200, 150, 120, 90],
    [110, 210, 160, 130, 95],
    [105, 205, 155, 125, 92],
    [115, 215, 165, 135, 100]
  ]
}
JSON

# send the request
response=$(curl -s -w "\nHTTP_STATUS:%{http_code}" -X POST "$BASE_URL/merge" \
    -H "Content-Type: application/json" \
    -d "$PAYLOAD")

# print response and status
echo "Response from server:"
echo "$response" | sed -n '1,/HTTP_STATUS:/p'
status=$(echo "$response" | sed -n 's/.*HTTP_STATUS://p')
echo "HTTP status: $status" 

if [[ "$status" == "200" ]]; then
    output=$(echo "$response" | grep -o '"outputPath"[^"]*"[^"]*"')
    echo "Merge output file reported: $output"
    echo "You can open or inspect that file in ./output"
else
    echo "Merge failed; check logs."
fi

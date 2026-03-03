#!/usr/bin/env bash
# start the app pointing at the benefits configuration and POST sample data

echo "Building and launching application with benefits config..."
cd /workspaces/codespaces-blank
export JAVA_HOME=/usr/lib/jvm/java-17-openjdk-amd64
export PATH=$JAVA_HOME/bin:$PATH
mvn clean package -DskipTests -q

# kill any existing server
lsof -ti:8080 | xargs kill -9 2>/dev/null || true

# start Spring Boot with custom config file location
java -jar -Dspring.config.location=application_benefits.yml target/excel-template-merger-1.0.0.jar &
APP_PID=$!

echo "Waiting for server to start..."
sleep 6

echo "Sending sample benefit JSON"
response=$(curl -s -w "\nHTTP_STATUS:%{http_code}" \
    -X POST http://localhost:8080/api/excel/merge \
    -H "Content-Type: application/json" \
    --data-binary @sample_benefits.json)

status=$(echo "$response" | sed -n 's/.*HTTP_STATUS://p')
echo "HTTP status: $status"
echo "$response" | sed -n '1,/HTTP_STATUS:/p'

echo "Output file should be in ./output. Inspect manually or with openpyxl."

echo "Shutting down app (pid $APP_PID)"
kill $APP_PID

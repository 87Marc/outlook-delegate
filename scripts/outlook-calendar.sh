#!/bin/bash
# Outlook Calendar Operations - Delegate Version
# Usage: outlook-calendar.sh <command> [args]
#
# This script accesses ANOTHER USER's calendar as a delegate.
# The assistant authenticates as itself, but reads/creates events on the owner's calendar.

CONFIG_DIR="$HOME/.outlook-mcp"
CREDS_FILE="$CONFIG_DIR/credentials.json"
CONFIG_FILE="$CONFIG_DIR/config.json"

# Load token
ACCESS_TOKEN=$(jq -r '.access_token' "$CREDS_FILE" 2>/dev/null)

if [ -z "$ACCESS_TOKEN" ] || [ "$ACCESS_TOKEN" = "null" ]; then
    echo '{"error": "No access token. Run setup first."}'
    exit 1
fi

# Load owner email (the calendar we're accessing as delegate)
OWNER_EMAIL=$(jq -r '.owner_email' "$CONFIG_FILE" 2>/dev/null)

if [ -z "$OWNER_EMAIL" ] || [ "$OWNER_EMAIL" = "null" ]; then
    echo '{"error": "No owner_email in config. Set the calendar owner in ~/.outlook-mcp/config.json"}'
    exit 1
fi

# Load timezone (default to UTC)
TIMEZONE=$(jq -r '.timezone // "UTC"' "$CONFIG_FILE" 2>/dev/null)

# DELEGATE ACCESS: Use /users/{owner} instead of /me
API="https://graph.microsoft.com/v1.0/users/$OWNER_EMAIL"

case "$1" in
    events)
        # List owner's upcoming events
        COUNT=${2:-10}
        curl -s "$API/calendar/events?\$top=$COUNT&\$orderby=start/dateTime%20desc&\$select=id,subject,start,end,location,isAllDay" \
            -H "Authorization: Bearer $ACCESS_TOKEN" \
            -H "Prefer: outlook.timezone=\"$TIMEZONE\"" | jq 'if .error then {error: .error.message} else (.value | to_entries | .[] | {n: (.key + 1), subject: .value.subject, start: .value.start.dateTime[0:16], end: .value.end.dateTime[0:16], location: (.value.location.displayName // ""), id: .value.id[-20:]}) end'
        ;;
    
    today)
        # List owner's today's events
        TODAY_START=$(date -u +"%Y-%m-%dT00:00:00Z")
        TODAY_END=$(date -u +"%Y-%m-%dT23:59:59Z")
        curl -s "$API/calendarView?startDateTime=$TODAY_START&endDateTime=$TODAY_END&\$orderby=start/dateTime&\$select=id,subject,start,end,location" \
            -H "Authorization: Bearer $ACCESS_TOKEN" \
            -H "Prefer: outlook.timezone=\"$TIMEZONE\"" | jq 'if .error then {error: .error.message} else (if .value then (.value | to_entries | .[] | {n: (.key + 1), subject: .value.subject, start: .value.start.dateTime[0:16], end: .value.end.dateTime[0:16], location: (.value.location.displayName // ""), id: .value.id[-20:]}) else {info: "No events today"} end) end'
        ;;
    
    week)
        # List owner's this week's events
        WEEK_START=$(date -u +"%Y-%m-%dT00:00:00Z")
        WEEK_END=$(date -u -d "+7 days" +"%Y-%m-%dT23:59:59Z" 2>/dev/null || date -u -v+7d +"%Y-%m-%dT23:59:59Z")
        curl -s "$API/calendarView?startDateTime=$WEEK_START&endDateTime=$WEEK_END&\$orderby=start/dateTime&\$select=id,subject,start,end,location,isAllDay" \
            -H "Authorization: Bearer $ACCESS_TOKEN" \
            -H "Prefer: outlook.timezone=\"$TIMEZONE\"" | jq 'if .error then {error: .error.message} else (if .value then (.value | to_entries | .[] | {n: (.key + 1), subject: .value.subject, start: .value.start.dateTime[0:16], end: .value.end.dateTime[0:16], location: (.value.location.displayName // ""), id: .value.id[-20:]}) else {info: "No events this week"} end) end'
        ;;
    
    read)
        # Read event details
        EVENT_ID="$2"
        FULL_ID=$(curl -s "$API/calendar/events?\$top=50&\$select=id" \
            -H "Authorization: Bearer $ACCESS_TOKEN" | jq -r ".value[] | select(.id | endswith(\"$EVENT_ID\")) | .id" | head -1)
        
        if [ -z "$FULL_ID" ]; then
            echo '{"error": "Event not found"}'
            exit 1
        fi
        
        curl -s "$API/calendar/events/$FULL_ID" \
            -H "Authorization: Bearer $ACCESS_TOKEN" \
            -H "Prefer: outlook.timezone=\"$TIMEZONE\"" | jq 'if .error then {error: .error.message} else {
                subject,
                start: .start.dateTime,
                end: .end.dateTime,
                location: .location.displayName,
                body: (if .body.contentType == "html" then (.body.content | gsub("<[^>]*>"; "") | gsub("\\s+"; " ")[0:500]) else .body.content[0:500] end),
                attendees: [.attendees[]?.emailAddress.address],
                isOnline: .isOnlineMeeting,
                link: .onlineMeeting.joinUrl,
                organizer: .organizer.emailAddress.address
            } end'
        ;;
    
    create)
        # Create event on owner's calendar
        SUBJECT="$2"
        START="$3"
        END="$4"
        LOCATION="${5:-}"
        
        if [ -z "$SUBJECT" ] || [ -z "$START" ] || [ -z "$END" ]; then
            echo "Usage: outlook-calendar.sh create <subject> <start> <end> [location]"
            echo "Date format: YYYY-MM-DDTHH:MM (e.g., 2026-01-26T10:00)"
            exit 1
        fi
        
        LOCATION_JSON=""
        if [ -n "$LOCATION" ]; then
            LOCATION_JSON=",\"location\": {\"displayName\": \"$LOCATION\"}"
        fi
        
        RESULT=$(curl -s "$API/calendar/events" \
            -H "Authorization: Bearer $ACCESS_TOKEN" \
            -H "Content-Type: application/json" \
            -d "{
                \"subject\": \"$SUBJECT\",
                \"start\": {\"dateTime\": \"$START\", \"timeZone\": \"$TIMEZONE\"},
                \"end\": {\"dateTime\": \"$END\", \"timeZone\": \"$TIMEZONE\"}
                $LOCATION_JSON
            }")
        
        echo "$RESULT" | jq 'if .error then {error: .error.message} else {status: "event created on '"$OWNER_EMAIL"' calendar", subject: .subject, start: .start.dateTime[0:16], end: .end.dateTime[0:16], id: .id[-20:]} end'
        ;;
    
    quick)
        # Quick event (1 hour from now or specified time)
        SUBJECT="$2"
        START_TIME="${3:-}"
        
        if [ -z "$SUBJECT" ]; then
            echo "Usage: outlook-calendar.sh quick <subject> [start-time]"
            echo "If no time given, creates 1-hour event starting now"
            exit 1
        fi
        
        if [ -z "$START_TIME" ]; then
            START=$(date +"%Y-%m-%dT%H:%M")
            END=$(date -d "+1 hour" +"%Y-%m-%dT%H:%M" 2>/dev/null || date -v+1H +"%Y-%m-%dT%H:%M")
        else
            START="$START_TIME"
            # Parse and add 1 hour
            END=$(date -d "$START_TIME + 1 hour" +"%Y-%m-%dT%H:%M" 2>/dev/null || echo "$START_TIME")
        fi
        
        RESULT=$(curl -s "$API/calendar/events" \
            -H "Authorization: Bearer $ACCESS_TOKEN" \
            -H "Content-Type: application/json" \
            -d "{
                \"subject\": \"$SUBJECT\",
                \"start\": {\"dateTime\": \"$START\", \"timeZone\": \"$TIMEZONE\"},
                \"end\": {\"dateTime\": \"$END\", \"timeZone\": \"$TIMEZONE\"}
            }")
        
        echo "$RESULT" | jq 'if .error then {error: .error.message} else {status: "quick event created on '"$OWNER_EMAIL"' calendar", subject: .subject, start: .start.dateTime[0:16], end: .end.dateTime[0:16], id: .id[-20:]} end'
        ;;
    
    delete)
        # Delete event from owner's calendar
        EVENT_ID="$2"
        FULL_ID=$(curl -s "$API/calendar/events?\$top=50&\$select=id" \
            -H "Authorization: Bearer $ACCESS_TOKEN" | jq -r ".value[] | select(.id | endswith(\"$EVENT_ID\")) | .id" | head -1)
        
        if [ -z "$FULL_ID" ]; then
            echo '{"error": "Event not found"}'
            exit 1
        fi
        
        RESULT=$(curl -s -w "\n%{http_code}" -X DELETE "$API/calendar/events/$FULL_ID" \
            -H "Authorization: Bearer $ACCESS_TOKEN")
        
        HTTP_CODE=$(echo "$RESULT" | tail -1)
        if [ "$HTTP_CODE" = "204" ]; then
            echo "{\"status\": \"event deleted from $OWNER_EMAIL calendar\", \"id\": \"$EVENT_ID\"}"
        else
            echo "$RESULT" | head -n -1 | jq '.error // .'
        fi
        ;;
    
    update)
        # Update event on owner's calendar
        EVENT_ID="$2"
        FIELD="$3"
        VALUE="$4"
        
        if [ -z "$FIELD" ] || [ -z "$VALUE" ]; then
            echo "Usage: outlook-calendar.sh update <id> <field> <value>"
            echo "Fields: subject, location, start, end"
            exit 1
        fi
        
        FULL_ID=$(curl -s "$API/calendar/events?\$top=50&\$select=id" \
            -H "Authorization: Bearer $ACCESS_TOKEN" | jq -r ".value[] | select(.id | endswith(\"$EVENT_ID\")) | .id" | head -1)
        
        if [ -z "$FULL_ID" ]; then
            echo '{"error": "Event not found"}'
            exit 1
        fi
        
        case "$FIELD" in
            subject)
                BODY="{\"subject\": \"$VALUE\"}"
                ;;
            location)
                BODY="{\"location\": {\"displayName\": \"$VALUE\"}}"
                ;;
            start)
                BODY="{\"start\": {\"dateTime\": \"$VALUE\", \"timeZone\": \"$TIMEZONE\"}}"
                ;;
            end)
                BODY="{\"end\": {\"dateTime\": \"$VALUE\", \"timeZone\": \"$TIMEZONE\"}}"
                ;;
            *)
                echo '{"error": "Unknown field: '"$FIELD"'", "valid_fields": ["subject", "location", "start", "end"]}'
                exit 1
                ;;
        esac
        
        curl -s -X PATCH "$API/calendar/events/$FULL_ID" \
            -H "Authorization: Bearer $ACCESS_TOKEN" \
            -H "Content-Type: application/json" \
            -d "$BODY" | jq 'if .error then {error: .error.message} else {status: "event updated", subject: .subject, start: .start.dateTime[0:16], end: .end.dateTime[0:16], id: .id[-20:]} end'
        ;;
    
    calendars)
        # List owner's calendars
        curl -s "$API/calendars" \
            -H "Authorization: Bearer $ACCESS_TOKEN" | jq 'if .error then {error: .error.message} else (.value[] | {name: .name, color: .color, canEdit: .canEdit, owner: .owner.address, id: .id[-20:]}) end'
        ;;
    
    free)
        # Check owner's free/busy for a time range
        START="$2"
        END="$3"
        
        if [ -z "$START" ] || [ -z "$END" ]; then
            echo "Usage: outlook-calendar.sh free <start> <end>"
            echo "Date format: YYYY-MM-DDTHH:MM"
            exit 1
        fi
        
        curl -s "$API/calendarView?startDateTime=${START}:00Z&endDateTime=${END}:00Z&\$select=subject,start,end" \
            -H "Authorization: Bearer $ACCESS_TOKEN" | jq 'if .error then {error: .error.message} else (if (.value | length) == 0 then {status: "free", owner: "'"$OWNER_EMAIL"'", start: "'"$START"'", end: "'"$END"'"} else {status: "busy", owner: "'"$OWNER_EMAIL"'", events: [.value[].subject]} end) end'
        ;;
    
    whoami)
        # Show delegate info
        DELEGATE=$(jq -r '.delegate_email // "unknown"' "$CONFIG_FILE" 2>/dev/null)
        echo "{\"delegate\": \"$DELEGATE\", \"accessing_calendar\": \"$OWNER_EMAIL\", \"timezone\": \"$TIMEZONE\", \"mode\": \"delegate\"}"
        ;;
    
    *)
        echo "Outlook Calendar - Delegate Access"
        echo "Accessing calendar: $OWNER_EMAIL"
        echo "Timezone: $TIMEZONE"
        echo ""
        echo "Usage: outlook-calendar.sh <command> [args]"
        echo ""
        echo "VIEW:"
        echo "  events [count]            - List upcoming events"
        echo "  today                     - Today's events"
        echo "  week                      - This week's events"
        echo "  read <id>                 - Event details"
        echo "  calendars                 - List all calendars"
        echo "  free <start> <end>        - Check availability"
        echo ""
        echo "CREATE (on $OWNER_EMAIL calendar):"
        echo "  create <subj> <start> <end> [loc] - Create event"
        echo "  quick <subject> [time]    - Quick 1-hour event"
        echo ""
        echo "MANAGE:"
        echo "  update <id> <field> <val> - Update event"
        echo "  delete <id>               - Delete event"
        echo ""
        echo "INFO:"
        echo "  whoami                    - Show delegate info"
        echo ""
        echo "Date format: YYYY-MM-DDTHH:MM (e.g., 2026-01-26T10:00)"
        ;;
esac

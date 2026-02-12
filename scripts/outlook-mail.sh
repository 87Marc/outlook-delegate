#!/bin/bash
# Outlook Mail Operations - Delegate Version
# Usage: outlook-mail.sh <command> [args]
#
# This script accesses ANOTHER USER's mailbox as a delegate.
# The assistant authenticates as itself, but reads/sends from the owner's mailbox.

CONFIG_DIR="$HOME/.outlook-mcp"
CREDS_FILE="$CONFIG_DIR/credentials.json"
CONFIG_FILE="$CONFIG_DIR/config.json"

# Load token
ACCESS_TOKEN=$(jq -r '.access_token' "$CREDS_FILE" 2>/dev/null)

if [ -z "$ACCESS_TOKEN" ] || [ "$ACCESS_TOKEN" = "null" ]; then
    echo '{"error": "No access token. Run setup first."}'
    exit 1
fi

# Load owner email (the mailbox we're accessing as delegate)
OWNER_EMAIL=$(jq -r '.owner_email' "$CONFIG_FILE" 2>/dev/null)

if [ -z "$OWNER_EMAIL" ] || [ "$OWNER_EMAIL" = "null" ]; then
    echo '{"error": "No owner_email in config. Set the mailbox owner in ~/.outlook-mcp/config.json"}'
    exit 1
fi

# DELEGATE ACCESS: Use /users/{owner} instead of /me
API="https://graph.microsoft.com/v1.0/users/$OWNER_EMAIL"

case "$1" in
    inbox)
        # List owner's inbox messages
        COUNT=${2:-10}
        curl -s "$API/messages?\$top=$COUNT&\$orderby=receivedDateTime%20desc&\$select=id,subject,from,receivedDateTime,isRead" \
            -H "Authorization: Bearer $ACCESS_TOKEN" | jq 'if .error then {error: .error.message} else (.value | to_entries | .[] | {n: (.key + 1), subject: .value.subject, from: .value.from.emailAddress.address, date: .value.receivedDateTime[0:16], read: .value.isRead, id: .value.id[-20:]}) end'
        ;;
    
    unread)
        # List owner's unread messages
        COUNT=${2:-20}
        curl -s "$API/messages?\$filter=isRead%20eq%20false&\$top=$COUNT&\$orderby=receivedDateTime%20desc&\$select=id,subject,from,receivedDateTime" \
            -H "Authorization: Bearer $ACCESS_TOKEN" | jq 'if .error then {error: .error.message} else (.value | to_entries | .[] | {n: (.key + 1), subject: .value.subject, from: .value.from.emailAddress.address, date: .value.receivedDateTime[0:16], id: .value.id[-20:]}) end'
        ;;
    
    search)
        # Search owner's emails
        QUERY="$2"
        COUNT=${3:-20}
        curl -s "$API/messages?\$search=\"$QUERY\"&\$top=$COUNT&\$select=id,subject,from,receivedDateTime" \
            -H "Authorization: Bearer $ACCESS_TOKEN" | jq 'if .error then {error: .error.message} else (.value | to_entries | .[] | {n: (.key + 1), subject: .value.subject, from: .value.from.emailAddress.address, date: .value.receivedDateTime[0:16], id: .value.id[-20:]}) end'
        ;;
    
    read)
        # Read specific email by ID (partial ID match - uses last 20 chars)
        MSG_ID="$2"
        # First find full ID (search by suffix)
        FULL_ID=$(curl -s "$API/messages?\$top=100&\$select=id" \
            -H "Authorization: Bearer $ACCESS_TOKEN" | jq -r ".value[] | select(.id | endswith(\"$MSG_ID\")) | .id" | head -1)
        
        if [ -z "$FULL_ID" ]; then
            echo '{"error": "Message not found. Use the ID shown in inbox/unread/search results."}'
            exit 1
        fi
        
        # Get message and extract text from HTML body
        curl -s "$API/messages/$FULL_ID?\$select=subject,from,receivedDateTime,body,toRecipients" \
            -H "Authorization: Bearer $ACCESS_TOKEN" | jq '{
                subject, 
                from: .from.emailAddress, 
                to: [.toRecipients[].emailAddress.address],
                date: .receivedDateTime,
                body: (if .body.contentType == "html" then (.body.content | gsub("<[^>]*>"; "") | gsub("\\s+"; " ") | gsub("&nbsp;"; " ") | .[0:2000]) else .body.content[0:2000] end)
            }'
        ;;
    
    mark-read)
        # Mark message as read
        MSG_ID="$2"
        FULL_ID=$(curl -s "$API/messages?\$top=100&\$select=id" \
            -H "Authorization: Bearer $ACCESS_TOKEN" | jq -r ".value[] | select(.id | endswith(\"$MSG_ID\")) | .id" | head -1)
        
        if [ -z "$FULL_ID" ]; then
            echo '{"error": "Message not found"}'
            exit 1
        fi
        
        curl -s -X PATCH "$API/messages/$FULL_ID" \
            -H "Authorization: Bearer $ACCESS_TOKEN" \
            -H "Content-Type: application/json" \
            -d '{"isRead": true}' | jq 'if .error then {error: .error.message} else {status: "marked as read", subject: .subject, id: .id[-20:]} end'
        ;;
    
    mark-unread)
        # Mark message as unread
        MSG_ID="$2"
        FULL_ID=$(curl -s "$API/messages?\$top=100&\$select=id" \
            -H "Authorization: Bearer $ACCESS_TOKEN" | jq -r ".value[] | select(.id | endswith(\"$MSG_ID\")) | .id" | head -1)
        
        if [ -z "$FULL_ID" ]; then
            echo '{"error": "Message not found"}'
            exit 1
        fi
        
        curl -s -X PATCH "$API/messages/$FULL_ID" \
            -H "Authorization: Bearer $ACCESS_TOKEN" \
            -H "Content-Type: application/json" \
            -d '{"isRead": false}' | jq 'if .error then {error: .error.message} else {status: "marked as unread", subject: .subject, id: .id[-20:]} end'
        ;;
    
    folders)
        # List owner's mail folders
        curl -s "$API/mailFolders" \
            -H "Authorization: Bearer $ACCESS_TOKEN" | jq 'if .error then {error: .error.message} else (.value[] | {name: .displayName, total: .totalItemCount, unread: .unreadItemCount}) end'
        ;;
    
    stats)
        # Get owner's inbox stats
        curl -s "$API/mailFolders/inbox" \
            -H "Authorization: Bearer $ACCESS_TOKEN" | jq 'if .error then {error: .error.message} else {folder: .displayName, total: .totalItemCount, unread: .unreadItemCount, owner: "'"$OWNER_EMAIL"'"} end'
        ;;
    
    send)
        # Send email ON BEHALF OF the owner
        # Recipient will see: "Assistant on behalf of Owner <owner@domain.com>"
        TO="$2"
        SUBJECT="$3"
        BODY="$4"
        
        if [ -z "$TO" ] || [ -z "$SUBJECT" ]; then
            echo 'Usage: outlook-mail.sh send <to> <subject> <body>'
            exit 1
        fi
        
        # DELEGATE SEND: Must specify 'from' as the owner
        # Using sendMail endpoint on the owner's mailbox
        RESULT=$(curl -s -w "\n%{http_code}" -X POST "$API/sendMail" \
            -H "Authorization: Bearer $ACCESS_TOKEN" \
            -H "Content-Type: application/json" \
            -d "{
                \"message\": {
                    \"subject\": \"$SUBJECT\",
                    \"body\": {\"contentType\": \"Text\", \"content\": \"$BODY\"},
                    \"toRecipients\": [{\"emailAddress\": {\"address\": \"$TO\"}}],
                    \"from\": {
                        \"emailAddress\": {
                            \"address\": \"$OWNER_EMAIL\"
                        }
                    }
                }
            }")
        
        HTTP_CODE=$(echo "$RESULT" | tail -1)
        if [ "$HTTP_CODE" = "202" ]; then
            echo "{\"status\": \"sent on behalf of $OWNER_EMAIL\", \"to\": \"$TO\", \"subject\": \"$SUBJECT\"}"
        else
            echo "$RESULT" | head -n -1 | jq '.error // .'
        fi
        ;;
    
    reply)
        # Reply to email on behalf of owner
        MSG_ID="$2"
        REPLY_BODY="$3"
        
        if [ -z "$MSG_ID" ] || [ -z "$REPLY_BODY" ]; then
            echo 'Usage: outlook-mail.sh reply <id> "reply body"'
            exit 1
        fi
        
        FULL_ID=$(curl -s "$API/messages?\$top=100&\$select=id" \
            -H "Authorization: Bearer $ACCESS_TOKEN" | jq -r ".value[] | select(.id | endswith(\"$MSG_ID\")) | .id" | head -1)
        
        if [ -z "$FULL_ID" ]; then
            echo '{"error": "Message not found"}'
            exit 1
        fi
        
        RESULT=$(curl -s -w "\n%{http_code}" -X POST "$API/messages/$FULL_ID/reply" \
            -H "Authorization: Bearer $ACCESS_TOKEN" \
            -H "Content-Type: application/json" \
            -d "{\"comment\": \"$REPLY_BODY\"}")
        
        HTTP_CODE=$(echo "$RESULT" | tail -1)
        if [ "$HTTP_CODE" = "202" ]; then
            echo "{\"status\": \"replied on behalf of $OWNER_EMAIL\", \"id\": \"$MSG_ID\"}"
        else
            echo "$RESULT" | head -n -1 | jq '.error // .'
        fi
        ;;
    
    flag)
        # Flag message as important
        MSG_ID="$2"
        FULL_ID=$(curl -s "$API/messages?\$top=100&\$select=id" \
            -H "Authorization: Bearer $ACCESS_TOKEN" | jq -r ".value[] | select(.id | endswith(\"$MSG_ID\")) | .id" | head -1)
        
        if [ -z "$FULL_ID" ]; then
            echo '{"error": "Message not found"}'
            exit 1
        fi
        
        curl -s -X PATCH "$API/messages/$FULL_ID" \
            -H "Authorization: Bearer $ACCESS_TOKEN" \
            -H "Content-Type: application/json" \
            -d '{"flag": {"flagStatus": "flagged"}}' | jq 'if .error then {error: .error.message} else {status: "flagged", subject: .subject, id: .id[-20:]} end'
        ;;
    
    unflag)
        # Remove flag
        MSG_ID="$2"
        FULL_ID=$(curl -s "$API/messages?\$top=100&\$select=id" \
            -H "Authorization: Bearer $ACCESS_TOKEN" | jq -r ".value[] | select(.id | endswith(\"$MSG_ID\")) | .id" | head -1)
        
        if [ -z "$FULL_ID" ]; then
            echo '{"error": "Message not found"}'
            exit 1
        fi
        
        curl -s -X PATCH "$API/messages/$FULL_ID" \
            -H "Authorization: Bearer $ACCESS_TOKEN" \
            -H "Content-Type: application/json" \
            -d '{"flag": {"flagStatus": "notFlagged"}}' | jq 'if .error then {error: .error.message} else {status: "unflagged", subject: .subject, id: .id[-20:]} end'
        ;;
    
    delete)
        # Move message to trash
        MSG_ID="$2"
        FULL_ID=$(curl -s "$API/messages?\$top=100&\$select=id" \
            -H "Authorization: Bearer $ACCESS_TOKEN" | jq -r ".value[] | select(.id | endswith(\"$MSG_ID\")) | .id" | head -1)
        
        if [ -z "$FULL_ID" ]; then
            echo '{"error": "Message not found"}'
            exit 1
        fi
        
        curl -s -X POST "$API/messages/$FULL_ID/move" \
            -H "Authorization: Bearer $ACCESS_TOKEN" \
            -H "Content-Type: application/json" \
            -d '{"destinationId": "deleteditems"}' | jq 'if .error then {error: .error.message} else {status: "moved to trash", subject: .subject, id: .id[-20:]} end'
        ;;
    
    archive)
        # Move message to archive
        MSG_ID="$2"
        FULL_ID=$(curl -s "$API/messages?\$top=100&\$select=id" \
            -H "Authorization: Bearer $ACCESS_TOKEN" | jq -r ".value[] | select(.id | endswith(\"$MSG_ID\")) | .id" | head -1)
        
        if [ -z "$FULL_ID" ]; then
            echo '{"error": "Message not found"}'
            exit 1
        fi
        
        curl -s -X POST "$API/messages/$FULL_ID/move" \
            -H "Authorization: Bearer $ACCESS_TOKEN" \
            -H "Content-Type: application/json" \
            -d '{"destinationId": "archive"}' | jq 'if .error then {error: .error.message} else {status: "archived", subject: .subject, id: .id[-20:]} end'
        ;;
    
    move)
        # Move message to folder
        MSG_ID="$2"
        FOLDER="$3"
        
        if [ -z "$FOLDER" ]; then
            echo 'Usage: outlook-mail.sh move <id> <folder-name>'
            exit 1
        fi
        
        FULL_ID=$(curl -s "$API/messages?\$top=100&\$select=id" \
            -H "Authorization: Bearer $ACCESS_TOKEN" | jq -r ".value[] | select(.id | endswith(\"$MSG_ID\")) | .id" | head -1)
        
        if [ -z "$FULL_ID" ]; then
            echo '{"error": "Message not found"}'
            exit 1
        fi
        
        # Find folder ID (case-insensitive)
        FOLDER_ID=$(curl -s "$API/mailFolders" \
            -H "Authorization: Bearer $ACCESS_TOKEN" | jq -r ".value[] | select(.displayName | ascii_downcase == \"$(echo "$FOLDER" | tr '[:upper:]' '[:lower:]')\") | .id" | head -1)
        
        if [ -z "$FOLDER_ID" ]; then
            echo '{"error": "Folder not found", "available": '$(curl -s "$API/mailFolders" -H "Authorization: Bearer $ACCESS_TOKEN" | jq '[.value[].displayName]')'}'
            exit 1
        fi
        
        curl -s -X POST "$API/messages/$FULL_ID/move" \
            -H "Authorization: Bearer $ACCESS_TOKEN" \
            -H "Content-Type: application/json" \
            -d "{\"destinationId\": \"$FOLDER_ID\"}" | jq 'if .error then {error: .error.message} else {status: "moved", folder: "'"$FOLDER"'", subject: .subject, id: .id[-20:]} end'
        ;;
    
    from)
        # List emails from specific sender
        SENDER="$2"
        COUNT=${3:-20}
        curl -s "$API/messages?\$filter=from/emailAddress/address%20eq%20'$SENDER'&\$top=$COUNT&\$orderby=receivedDateTime%20desc&\$select=id,subject,from,receivedDateTime" \
            -H "Authorization: Bearer $ACCESS_TOKEN" | jq 'if .error then {error: .error.message} else (.value | to_entries | .[] | {n: (.key + 1), subject: .value.subject, from: .value.from.emailAddress.address, date: .value.receivedDateTime[0:16], id: .value.id[-20:]}) end'
        ;;
    
    attachments)
        # List attachments
        MSG_ID="$2"
        FULL_ID=$(curl -s "$API/messages?\$top=100&\$select=id" \
            -H "Authorization: Bearer $ACCESS_TOKEN" | jq -r ".value[] | select(.id | endswith(\"$MSG_ID\")) | .id" | head -1)
        
        if [ -z "$FULL_ID" ]; then
            echo '{"error": "Message not found"}'
            exit 1
        fi
        
        curl -s "$API/messages/$FULL_ID/attachments" \
            -H "Authorization: Bearer $ACCESS_TOKEN" | jq 'if .error then {error: .error.message} else (.value | to_entries | .[] | {n: (.key + 1), name: .value.name, size: .value.size, type: .value.contentType, id: .value.id[-20:]}) end'
        ;;
    
    whoami)
        # Show delegate info - who is accessing whose mailbox
        DELEGATE=$(jq -r '.delegate_email // "unknown"' "$CONFIG_FILE" 2>/dev/null)
        echo "{\"delegate\": \"$DELEGATE\", \"accessing_mailbox\": \"$OWNER_EMAIL\", \"mode\": \"delegate\"}"
        ;;
    
    *)
        echo "Outlook Mail - Delegate Access"
        echo "Accessing mailbox: $OWNER_EMAIL"
        echo ""
        echo "Usage: outlook-mail.sh <command> [args]"
        echo ""
        echo "READING:"
        echo "  inbox [count]             - List latest emails"
        echo "  unread [count]            - List unread emails"
        echo "  search \"query\" [count]    - Search emails"
        echo "  from <email> [count]      - Emails from sender"
        echo "  read <id>                 - Read email content"
        echo "  attachments <id>          - List attachments"
        echo ""
        echo "MANAGING:"
        echo "  mark-read <id>            - Mark as read"
        echo "  mark-unread <id>          - Mark as unread"
        echo "  flag <id>                 - Flag as important"
        echo "  unflag <id>               - Remove flag"
        echo "  delete <id>               - Move to trash"
        echo "  archive <id>              - Move to archive"
        echo "  move <id> <folder>        - Move to folder"
        echo ""
        echo "SENDING (on behalf of $OWNER_EMAIL):"
        echo "  send <to> <subj> <body>   - Send new email"
        echo "  reply <id> \"body\"         - Reply to email"
        echo ""
        echo "INFO:"
        echo "  folders                   - List mail folders"
        echo "  stats                     - Inbox statistics"
        echo "  whoami                    - Show delegate info"
        ;;
esac

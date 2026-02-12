#!/bin/bash
# Outlook Token Management - Delegate Version
# Usage: outlook-token.sh <command>

CONFIG_DIR="$HOME/.outlook-mcp"
CONFIG_FILE="$CONFIG_DIR/config.json"
CREDS_FILE="$CONFIG_DIR/credentials.json"

# Load config
CLIENT_ID=$(jq -r '.client_id' "$CONFIG_FILE" 2>/dev/null)
CLIENT_SECRET=$(jq -r '.client_secret' "$CONFIG_FILE" 2>/dev/null)
OWNER_EMAIL=$(jq -r '.owner_email' "$CONFIG_FILE" 2>/dev/null)
DELEGATE_EMAIL=$(jq -r '.delegate_email' "$CONFIG_FILE" 2>/dev/null)
REFRESH_TOKEN=$(jq -r '.refresh_token' "$CREDS_FILE" 2>/dev/null)

case "$1" in
    refresh)
        if [ -z "$REFRESH_TOKEN" ] || [ "$REFRESH_TOKEN" = "null" ]; then
            echo '{"error": "No refresh token. Run setup first."}'
            exit 1
        fi
        
        RESPONSE=$(curl -s -X POST "https://login.microsoftonline.com/common/oauth2/v2.0/token" \
            -d "client_id=$CLIENT_ID" \
            -d "client_secret=$CLIENT_SECRET" \
            -d "refresh_token=$REFRESH_TOKEN" \
            -d "grant_type=refresh_token" \
            -d "scope=offline_access%20User.Read%20Mail.ReadWrite.Shared%20Mail.Send.Shared%20Calendars.ReadWrite.Shared")
        
        if echo "$RESPONSE" | jq -e '.access_token' > /dev/null 2>&1; then
            echo "$RESPONSE" > "$CREDS_FILE"
            echo '{"status": "token refreshed", "expires_in": '$(echo "$RESPONSE" | jq '.expires_in')'}'
        else
            echo "$RESPONSE" | jq '.error_description // .error // .'
        fi
        ;;
    
    test)
        ACCESS_TOKEN=$(jq -r '.access_token' "$CREDS_FILE" 2>/dev/null)
        
        if [ -z "$ACCESS_TOKEN" ] || [ "$ACCESS_TOKEN" = "null" ]; then
            echo '{"error": "No access token"}'
            exit 1
        fi
        
        echo "Testing delegate access..."
        echo ""
        
        # Test 1: Delegate's own identity
        echo "1. Delegate identity (who is authenticated):"
        DELEGATE_INFO=$(curl -s "https://graph.microsoft.com/v1.0/me" \
            -H "Authorization: Bearer $ACCESS_TOKEN")
        echo "$DELEGATE_INFO" | jq '{authenticated_as: .userPrincipalName, display_name: .displayName}'
        
        echo ""
        
        # Test 2: Access to owner's mailbox
        echo "2. Owner mailbox access ($OWNER_EMAIL):"
        OWNER_INBOX=$(curl -s "https://graph.microsoft.com/v1.0/users/$OWNER_EMAIL/mailFolders/inbox" \
            -H "Authorization: Bearer $ACCESS_TOKEN")
        
        if echo "$OWNER_INBOX" | jq -e '.error' > /dev/null 2>&1; then
            echo "$OWNER_INBOX" | jq '{error: .error.message, code: .error.code}'
            echo ""
            echo "⚠️  Cannot access owner's mailbox. Check delegate permissions."
        else
            echo "$OWNER_INBOX" | jq '{status: "OK", folder: .displayName, unread: .unreadItemCount, total: .totalItemCount}'
        fi
        
        echo ""
        
        # Test 3: Access to owner's calendar
        echo "3. Owner calendar access ($OWNER_EMAIL):"
        OWNER_CAL=$(curl -s "https://graph.microsoft.com/v1.0/users/$OWNER_EMAIL/calendar" \
            -H "Authorization: Bearer $ACCESS_TOKEN")
        
        if echo "$OWNER_CAL" | jq -e '.error' > /dev/null 2>&1; then
            echo "$OWNER_CAL" | jq '{error: .error.message, code: .error.code}'
            echo ""
            echo "⚠️  Cannot access owner's calendar. Check delegate permissions."
        else
            echo "$OWNER_CAL" | jq '{status: "OK", calendar: .name, canEdit: .canEdit}'
        fi
        
        echo ""
        echo "Summary:"
        echo "  Delegate: $DELEGATE_EMAIL"
        echo "  Owner:    $OWNER_EMAIL"
        echo "  Mode:     Delegate Access"
        ;;
    
    get)
        ACCESS_TOKEN=$(jq -r '.access_token' "$CREDS_FILE" 2>/dev/null)
        if [ -z "$ACCESS_TOKEN" ] || [ "$ACCESS_TOKEN" = "null" ]; then
            echo '{"error": "No access token"}'
            exit 1
        fi
        echo "$ACCESS_TOKEN"
        ;;
    
    info)
        echo "Delegate Configuration:"
        echo "  Config dir:    $CONFIG_DIR"
        echo "  Delegate:      $DELEGATE_EMAIL"
        echo "  Owner:         $OWNER_EMAIL"
        echo "  Client ID:     ${CLIENT_ID:0:8}..."
        
        if [ -f "$CREDS_FILE" ]; then
            EXPIRES=$(jq -r '.expires_in // "unknown"' "$CREDS_FILE")
            echo "  Token exists:  yes"
            echo "  Expires in:    $EXPIRES seconds (from last refresh)"
        else
            echo "  Token exists:  no"
        fi
        ;;
    
    *)
        echo "Outlook Token Management - Delegate Version"
        echo ""
        echo "Usage: outlook-token.sh <command>"
        echo ""
        echo "Commands:"
        echo "  refresh  - Refresh access token"
        echo "  test     - Test connection to BOTH delegate and owner accounts"
        echo "  get      - Print current access token"
        echo "  info     - Show configuration info"
        ;;
esac

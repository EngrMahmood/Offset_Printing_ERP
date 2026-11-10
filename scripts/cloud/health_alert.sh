#!/usr/bin/env bash
# Health alert for the Offset ERP servers. Runs from cron on BOTH the primary and the standby, so if one
# server is down the other one notices and emails.
#
# Checks (any failure becomes an email to the administrator):
#   * every site answers over HTTPS with a valid certificate (a dead or expired site shows as unreachable)
#   * the certificate of each site has more than 21 days left (renewal normally happens at 30 days, so
#     reaching 21 means automatic renewal is failing)
#   * the web, nginx and redis containers on THIS host are running
#   * the disk is below 90 percent full
#
# Mail goes through the ERP's own Gmail account (manage.py send_ops_alert). The same problem is mailed again at
# most every 12 hours, and a "resolved" mail is sent when it clears.
#
# Install (as the ubuntu user):   */30 * * * * /home/ubuntu/offset-erp/scripts/cloud/health_alert.sh
set -u

APP_DIR="${APP_DIR:-$HOME/offset-erp}"
SITES="${ERP_SITES:-offseterp.duckdns.org offseterpbackup.duckdns.org}"
WARN_DAYS="${WARN_DAYS:-21}"
CRIT_DAYS="${CRIT_DAYS:-7}"
DISK_LIMIT="${DISK_LIMIT:-90}"
REPEAT_SECONDS="${REPEAT_SECONDS:-43200}"
STATE_DIR="$HOME/.erp_health_alert"
HOST_NAME="$(hostname)"
mkdir -p "$STATE_DIR"
LOG="$STATE_DIR/alert.log"

issues=()

now_epoch=$(date +%s)

for site in $SITES; do
    code=$(curl -s -m 20 -o /dev/null -w '%{http_code}' "https://$site/login/" 2>/dev/null)
    case "$code" in
        2*|3*|4*) ;;                                       # the site answered
        *) issues+=("UNREACHABLE: https://$site did not answer properly (HTTP status $code). Browsers will show an error if its certificate expired.") ;;
    esac

    end=$(echo | timeout 20 openssl s_client -connect "$site:443" -servername "$site" 2>/dev/null | openssl x509 -noout -enddate 2>/dev/null | cut -d= -f2)
    if [ -n "$end" ]; then
        end_epoch=$(date -d "$end" +%s 2>/dev/null || echo 0)
        days=$(( (end_epoch - now_epoch) / 86400 ))
        if [ "$end_epoch" -le "$now_epoch" ]; then
            issues+=("CERTIFICATE EXPIRED: $site expired on $end.")
        elif [ "$days" -le "$CRIT_DAYS" ]; then
            issues+=("CERTIFICATE CRITICAL: $site expires in $days day(s) ($end). Automatic renewal is not working.")
        elif [ "$days" -le "$WARN_DAYS" ]; then
            issues+=("CERTIFICATE WARNING: $site expires in $days days ($end). Automatic renewal should already have happened.")
        fi
    else
        issues+=("CERTIFICATE UNREADABLE: could not read the certificate of $site.")
    fi
done

running=$(docker ps --format '{{.Names}}' 2>/dev/null)
for part in web nginx redis; do
    echo "$running" | grep -q "offset-erp-$part-1" || issues+=("CONTAINER DOWN on $HOST_NAME: offset-erp-$part-1 is not running.")
done

used=$(df -P / | awk 'NR==2 {gsub("%","",$5); print $5}')
if [ -n "$used" ] && [ "$used" -ge "$DISK_LIMIT" ]; then
    issues+=("DISK FULL on $HOST_NAME: root disk is ${used}% used.")
fi

send_mail() {   # subject, body
    ( cd "$APP_DIR" && docker compose exec -T web python manage.py send_ops_alert --subject "$1" --body "$2" ) >>"$LOG" 2>&1
}

signature=$(printf '%s\n' "${issues[@]:-}" | sort | md5sum | cut -d' ' -f1)
sig_file="$STATE_DIR/last_signature"
time_file="$STATE_DIR/last_sent"
last_sig=$(cat "$sig_file" 2>/dev/null || echo "")
last_sent=$(cat "$time_file" 2>/dev/null || echo 0)

if [ "${#issues[@]}" -eq 0 ]; then
    if [ -n "$last_sig" ]; then
        send_mail "[Offset ERP] RESOLVED - all checks pass ($HOST_NAME)" "All checks pass again on $(date '+%Y-%m-%d %H:%M %Z') (seen from $HOST_NAME)."
        rm -f "$sig_file" "$time_file"
    fi
    exit 0
fi

if [ "$signature" != "$last_sig" ] || [ $(( now_epoch - last_sent )) -ge "$REPEAT_SECONDS" ]; then
    body="Problems found by $HOST_NAME at $(date '+%Y-%m-%d %H:%M %Z'):

$(printf ' - %s\n' "${issues[@]}")

Primary:  https://offseterp.duckdns.org
Standby:  https://offseterpbackup.duckdns.org
Certificate renewal log on the server: sudo tail -n 50 /var/log/letsencrypt/letsencrypt.log"
    if send_mail "[Offset ERP] ALERT from $HOST_NAME: ${#issues[@]} problem(s) - ${issues[0]:0:70}" "$body"; then
        echo "$signature" >"$sig_file"
        echo "$now_epoch" >"$time_file"
    else
        echo "$(date) alert mail could not be sent (is the web container up?)" >>"$LOG"
    fi
fi
exit 0

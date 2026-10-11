#!/bin/bash
set -euo pipefail

if [[ "$#" -ne 8 ]]; then
  printf 'usage: %s plist node server working-dir asset-dir cert key version\n' "$0" >&2
  exit 64
fi

plist_path="$1"
node_path="$2"
server_path="$3"
working_dir="$4"
asset_dir="$5"
cert_file="$6"
key_file="$7"
version="$8"
label="com.analystlens.local-host"

xml_escape() {
  local value="$1"
  value="${value//&/&amp;}"
  value="${value//</&lt;}"
  value="${value//>/&gt;}"
  value="${value//\"/&quot;}"
  value="${value//\'/&apos;}"
  printf '%s' "$value"
}

mkdir -p "$(dirname "$plist_path")" "${working_dir}/logs"
temporary_plist="${plist_path}.tmp.$$"
trap 'rm -f "$temporary_plist"' EXIT
umask 077

{
  printf '%s\n' '<?xml version="1.0" encoding="UTF-8"?>'
  printf '%s\n' '<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">'
  printf '%s\n' '<plist version="1.0">'
  printf '%s\n' '<dict>'
  printf '%s\n' '  <key>Label</key>'
  printf '  <string>%s</string>\n' "$(xml_escape "$label")"
  printf '%s\n' '  <key>ProgramArguments</key>'
  printf '%s\n' '  <array>'
  printf '    <string>%s</string>\n' "$(xml_escape "$node_path")"
  printf '    <string>%s</string>\n' "$(xml_escape "$server_path")"
  printf '%s\n' '  </array>'
  printf '%s\n' '  <key>EnvironmentVariables</key>'
  printf '%s\n' '  <dict>'
  printf '%s\n' '    <key>ANALYST_LENS_ASSET_DIR</key>'
  printf '    <string>%s</string>\n' "$(xml_escape "$asset_dir")"
  printf '%s\n' '    <key>ANALYST_LENS_CERT_FILE</key>'
  printf '    <string>%s</string>\n' "$(xml_escape "$cert_file")"
  printf '%s\n' '    <key>ANALYST_LENS_KEY_FILE</key>'
  printf '    <string>%s</string>\n' "$(xml_escape "$key_file")"
  printf '%s\n' '    <key>ANALYST_LENS_VERSION</key>'
  printf '    <string>%s</string>\n' "$(xml_escape "$version")"
  printf '%s\n' '  </dict>'
  printf '%s\n' '  <key>WorkingDirectory</key>'
  printf '  <string>%s</string>\n' "$(xml_escape "$working_dir")"
  printf '%s\n' '  <key>RunAtLoad</key>'
  printf '%s\n' '  <true/>'
  printf '%s\n' '  <key>KeepAlive</key>'
  printf '%s\n' '  <true/>'
  printf '%s\n' '  <key>ThrottleInterval</key>'
  printf '%s\n' '  <integer>3</integer>'
  printf '%s\n' '  <key>StandardOutPath</key>'
  printf '  <string>%s</string>\n' "$(xml_escape "${working_dir}/logs/host.log")"
  printf '%s\n' '  <key>StandardErrorPath</key>'
  printf '  <string>%s</string>\n' "$(xml_escape "${working_dir}/logs/host-error.log")"
  printf '%s\n' '</dict>'
  printf '%s\n' '</plist>'
} > "$temporary_plist"

chmod 600 "$temporary_plist"
mv "$temporary_plist" "$plist_path"
trap - EXIT

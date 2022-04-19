#! /bin/bash

function ok_echo {
  echo $(tput setaf 2)"$1"$(tput init)
}

function error_echo {
  echo $(tput setaf 1)"$1"$(tput init)
}

function fatal_error {
  error_echo "$1"
  exit 1
}

app="$1"
test ! -z "$app" || fatal_error "usage: $0 app-directory"
test "${app: -4}" == ".app" -a -d "$app/Contents" || fatal_error "$app: not an Application folder"

if ! /usr/sbin/DevToolsSecurity | grep -q enabled; then
  error_echo "Developer mode disabled; lldb will fail to attach. Enable with:"
  error_echo "    sudo DevToolsSecurity -enable"
  exit 1
fi

if ! /usr/bin/csrutil status | grep -q disabled; then
  error_echo "System Integrity Protection status enabled; lldb will fail to attach. Enable by:"
  error_echo "    1. Enter Recover mode (boot with Command-R / hold power button)"
  error_echo "    2. Start Terminal"
  error_echo "    3. csrutil disable"
  error_echo "    4. reboot"
  exit
fi

ok_echo "Opening $app"
open --hide -a "$app"
pid=`/bin/ps auxwww | awk '/Numbers$/{print $2}'`

ok_echo "Waiting a while (5s)"
sleep 5

ok_echo "Dumping in lldb"
echo 'po [TSPRegistry sharedRegistry]' | /usr/bin/lldb -p $pid > protos/TSPRegistry.dump

kill -HUP $pid

ok_echo "Dumped mapping to protos/TSPRegistry.dump"

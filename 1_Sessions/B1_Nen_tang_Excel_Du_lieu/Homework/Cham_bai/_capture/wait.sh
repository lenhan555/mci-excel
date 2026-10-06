#!/bin/bash
# wait.sh NNN_name  -> chờ log có DONE
for i in $(seq 1 80); do
  if [ -f "out/$1.log" ] && grep -q DONE "out/$1.log"; then sed 's/^\xEF\xBB\xBF//' "out/$1.log"; exit 0; fi
  sleep 2
done
echo "TIMEOUT; cmd còn: $(ls cmd)"

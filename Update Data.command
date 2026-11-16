#!/bin/bash
# Run from this file's folder, even after the folder has been moved.
MORTGAGE_DIR="$(cd -- "$(dirname -- "$0")" && pwd)" || exit 1
cd "$MORTGAGE_DIR" || exit 1
MORTGAGE_PY=""
for MORTGAGE_CANDIDATE in python3 python "$HOME/anaconda3/bin/python" "$HOME/miniconda3/bin/python" /opt/anaconda3/bin/python /opt/miniconda3/bin/python /opt/homebrew/bin/python3 /usr/local/bin/python3; do
  if "$MORTGAGE_CANDIDATE" -c 'import sys; sys.exit(sys.version_info < (3, 10))' >/dev/null 2>&1; then
    MORTGAGE_PY="$MORTGAGE_CANDIDATE"
    break
  fi
done
if [ -z "$MORTGAGE_PY" ]; then
  echo "Updating requires Python 3.10 or newer. Install Python from python.org, then run this file again."
  echo "You can still open HK_Mortgage_Dashboard.html without Python or internet."
  MORTGAGE_EXIT=1
else
  "$MORTGAGE_PY" "$MORTGAGE_DIR/update_mortgage_data.py" "$@"
  MORTGAGE_EXIT=$?
  if [ "$MORTGAGE_EXIT" -eq 0 ]; then
    echo "Completed. Open HK_Mortgage_Dashboard.html or reload its browser tab."
  else
    echo "Update did not finish. See the message above and update_status.txt if present."
  fi
fi
if [ -t 0 ]; then
  read -r -p "Press Return to close. " MORTGAGE_UNUSED
fi
exit "$MORTGAGE_EXIT"

#!/bin/bash
MORTGAGE_DIR="$(cd -- "$(dirname -- "$0")" && pwd)" || exit 1
open "$MORTGAGE_DIR/HK_Mortgage_Dashboard.html"

#!/bin/bash

SCRIPT_DIR=$(cd $(dirname $0); pwd)
MAC_DIR="$SCRIPT_DIR/macos-xlsx2csv"
mkdir -p $MAC_DIR

set -e
cargo build --release
cp target/release/xlsx2csv $MAC_DIR/
cp README.md $MAC_DIR/
cp README-*.md $MAC_DIR/
cp LICENSE $MAC_DIR/
cp Cargo.toml $MAC_DIR/
cp xlsx2csv.toml $MAC_DIR/
cd $MAC_DIR && zip $SCRIPT_DIR/macos-xlsx2csv.zip *

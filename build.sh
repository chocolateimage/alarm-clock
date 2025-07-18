#!/usr/bin/env bash

VERSION=1.4.0-1

DIRECTORY_NAME="alarm-clock-$VERSION"

rm -rf debian_dist
mkdir -p debian_dist
cd debian_dist
mkdir -p "$DIRECTORY_NAME/DEBIAN"
cp ../control "$DIRECTORY_NAME/DEBIAN/"
mkdir -p "$DIRECTORY_NAME/usr/bin"
cp ../alarm-clock.py "$DIRECTORY_NAME/usr/bin/alarm-clock"
mkdir -p "$DIRECTORY_NAME/usr/share/applications"
cp ../alarm-clock.desktop "$DIRECTORY_NAME/usr/share/applications/"

dpkg-deb --build "$DIRECTORY_NAME"
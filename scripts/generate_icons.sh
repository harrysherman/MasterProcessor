#!/bin/bash

# Installer Script for new systems
unameOut="$(uname -s)"

case "${unameOut}" in
    Linux*)     machine=Linux;;
    Darwin*)    machine=Mac;;
    CYGWIN*)    machine=Cygwin;;
    MINGW*)     machine=MinGw;;
    *)          machine="UNKNOWN:${unameOut}"
esac

PNG_MASTER="icons/AG_circle_rgb_gold.png"

ICONSET_FOLDER="AppIcon.iconset"
sizes=(
  16x16
  32x32
  128x128
  256x256
  512x512
)

# Generate renditions at the sizes in "sizes" above, put all in ICONSET_FOLDER
mkdir -p $ICONSET_FOLDER
for size in "${sizes[@]}"; do
  icon="icon_${size}.png"
  ICON_FILES="$ICON_FILES $ICONSET_FOLDER/$icon"
  echo Generating $ICONSET_FOLDER/"$icon"
  # convert $PNG_MASTER -quality 100 -resize $size $ICONSET_FOLDER/$icon
  sips -z "$size" "$size" $PNG_MASTER --out $ICONSET_FOLDER/"$icon"
  
  icon="icon_${size}@2x.png"
  ICON_FILES="$ICON_FILES $ICONSET_FOLDER/$icon"
  echo Generating $ICONSET_FOLDER/"$icon"
  # convert $PNG_MASTER -quality 100 -resize $size $ICONSET_FOLDER/$icon
  sips -z "$size" "$size" $PNG_MASTER --out $ICONSET_FOLDER/"$icon"
done

if [[ $machine == "Mac" ]]
then
# generate icon.icns for mac app (this only works on mac)
echo Generating icon.icns
iconutil -c icns $ICONSET_FOLDER -o icons/icon.icns
else

# Generate .ico file for windows
ICON_FILES=""
for size in "${sizes[@]}"; do
  ICON_FILES="$ICON_FILES $ICONSET_FOLDER/icon_${size}.png"
  ICON_FILES="$ICON_FILES $ICONSET_FOLDER/icon_${size}@2x.png"
done
echo Generating icon.ico
convert $ICON_FILES icons/icon.ico 
fi
# remove generated renditions
echo removing $ICONSET_FOLDER folder
rm -rf $ICONSET_FOLDER
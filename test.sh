#!/bin/bash

for D in dashboard-{1..12}; do
  echo 'testing' $D
  cd $D
  ./main.py dummy-data.xlsx && rm output.pptx
  cd ..
done

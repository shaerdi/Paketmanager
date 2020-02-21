#!/bin/bash

for i in $(ls *ui)
do
    pyuic5 $i -o ${i/.ui/.py}
done

for i in $(ls *qrc)
do
    pyrcc5 $i -o ${i/.qrc/_rc.py}
done

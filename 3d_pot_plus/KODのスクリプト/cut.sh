#!/bin/sh
#切り取って整数化するスクリプト
RESULT = cut -d " " -f 1|SEISUU='echo"scale;$RESULT * 10000" |bc'
echo $?

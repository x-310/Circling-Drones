#!/bin/sh
#高さzを抜き出す
sed -e 's/ \+/ /g'|sed -e '1,2d'|cut -f 5 -d " "
exit $?

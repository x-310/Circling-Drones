#!/bin/sh
#スペースをすべて削除のちスペース一つ追加/1,2行目を削除/文字E+04とE+03を削除/3,4行目を切り出す
sed -e 's/ \+/ /g'|sed -e '1,2d'|sed -e 's/E+04//g'|sed -e 's/E+03//g'| cut -f 3,4 -d " "
exit $?

//地形部分を表すcsvを作るためのプログラム
#include <stdio.h>
#include <stdlib.h>

int main(){
 FILE *fp;
 FILE *foutp;
  float takasa;

 fp = fopen("sabuntakasa.txt","r");
 foutp = fopen("sabuntakasa.csv","w");
//エラー処理
 if(fp == NULL){
	printf("FILE OPEN ERROR\n");
	return -1;
 }

 if(foutp == NULL){
	printf("FILE OPEN ERROR\n");
	return -1;
 }

//メートル表記の入る変数
 int tx;
 int ty;

//csvに座標を入れる処理
 for(int j = 0 ;j < 52 ;j++){
	for(int i = 0 ;i < 94 ;i++){
		tx = i * 5;
		ty = j * 5;
		fscanf(fp,"%f",&takasa);
		fprintf(foutp, "%d,%d,%f\n",tx,ty,takasa);
		}
	}

 fclose(fp);
 fclose(foutp);

 return 0;
 }

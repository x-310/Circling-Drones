//ファイルからMIN_HEIGHTを引くプログラム
#include <stdio.h>
#include <stdlib.h>


#define MIN_HEIGHT 1542.969//min_search.cの結果

int main(){
 FILE *fp;
 FILE *foutp;
 int ret;
 float takasa;

 fp = fopen("takasa.txt","r");
 foutp = fopen("sabun.csv","w");

//エラー処理
 if(fp == NULL){
	printf("FILE OPEN ERROR\n");
	return -1;
 }

 if(foutp == NULL){
	printf("FILE OPEN ERROR\n");
	return -1;
 }

 while((ret = fscanf(fp,"%f",&takasa))
		!= EOF){
	takasa = takasa - MIN_HEIGHT;
	fprintf(foutp,"%f\n",takasa);
 }

 fclose(fp);
 fclose(foutp);

 return 0;
 }

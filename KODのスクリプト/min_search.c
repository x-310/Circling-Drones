//ファイルから最小値を探すプログラム
#include <stdio.h>
#include <stdlib.h>

#define MIN 65535
#define MAX 0
int main(){
 FILE *fp;
 int ret;
 float takasa;
 float min = MIN;
 //float max = MAX;
 fp = fopen("takasa.txt","r");//各座標の高さのみが入ったファイル

 if(fp == NULL){
	printf("FLLE OPEN ERROR!!\n");
	return -1;
 }

 while((ret = fscanf(fp,"%f",&takasa))
			!= EOF){

 if(takasa < min){
	min = takasa;
	}
 }
//最小値を表示
 printf("最小値は%f\n",min);

//  if(takasa > max){
//	max = takasa;
//	}
//  }

//  printf("最大値は%f\n",max);
 fclose(fp);

 return 0;
}

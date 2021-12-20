//一定高さでの受信機電力が書かれたファイルに座標を与えるプログラム
#include <fstream>
#include <iostream>
#include <string>
#include <vector>
#include <stdlib.h>
#include <math.h>
#include <cmath>
#include "float.h"

using namespace std;

#define MAPSIZE_X 94
#define MAPSIZE_Y 52
#define MAPSIZE_Z 9
#define TRANSMITTER_OUTPUT 23//送信機の出力

//ファイルの名前用
string filename_parts1 = "d_power.txt";//受信電力のみが入ったファイル
string filename_parts2 = "mreceived_power.csv";//地表の形を考慮した受信電力ファイル
string filename_parts3 = "greceived_power.csv";//平らな形の受信電力ファイル

int main(){
	//地形による高さ変化
	ifstream ifs("sabun.csv");
	//エラー文
	if(!ifs){
		cout << "Could not input sabun.csv" << endl;
		return -1;
	}
	else cout << "sabun.csv opened" << endl;
	string str;
	vector<double> input;

	while(getline(ifs,str)){
		input.push_back(atof(str.c_str()));
	}
	/////////////////////////////////////////////
	//受信電力部分のファイル
	//string str_t;
	//vector<double> input_t;

	for(int height = 10; height <= 50; height = height + 5){
		string str2;
	        vector<double> input2;

		string height_s = to_string(height);
		string file_name1 = "h" + height_s + "d_power.txt";
		string file_name2 = height_s + filename_parts2;
		string file_name3 = height_s + filename_parts3;
		ifstream ifs_t(file_name1);
		if(!ifs_t){
			cout << "creative " << file_name1 << " could not be open" << endl;
			return -1;
			}
		else cout << "creative " << file_name1 << " opend" << endl;

		while(getline(ifs_t,str2)){
			input2.push_back(atof(str2.c_str()));
		}

		ofstream fout(file_name2);
		if(!fout){
			cout << "creative " << file_name2 << " could not be open" << endl;
			return -1;
		}
		else cout << "creative " << file_name2 << " opened" << endl;

		ofstream fout_t(file_name3);
                if(!fout_t){
                        cout << "creative " << file_name3 << " could not be open" << endl;
                        return -1;
                }
                else cout << "creative " << file_name3 << " opened" << endl;
		int s = 0;
		int x;
		int y;
 		int z;
		int tx;
		int ty;
		float tz;
			for(int j = 0;j < MAPSIZE_Y; j++){
				for(int i = 0;i < MAPSIZE_X; i++){
					if(input2[s]){
				 	input2[s] = input2[s];
						if(input[s]){
							input[s] = input[s];

							//受信電力が送信出力を超えるときは送信出力の値とする
							if(input2[s] > TRANSMITTER_OUTPUT){
								input2[s] = TRANSMITTER_OUTPUT;
								}
							//////////////////////

							x = i;
							y = j;
							z = height;
							tx = i * 5;
							ty = j * 5;
							//cout << inputt[s] << endl;
							tz = height + input[s]; //パワーの高さ+地形高さ
							fout << tx << "," << ty << "," << tz << "," << input2[s] << endl;
							fout_t << tx << "," << ty << "," << z << "," << input2[s] << endl;
							s++;
							}
						}
					}
				}
		fout.close();
		fout_t.close();
		}
	return 0;
 }

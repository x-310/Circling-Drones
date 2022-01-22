//ポテンシャル法で3次元の経路を求めるプログラム
//2020/10/2　弱い受信機電力を避けていく経路を求める
//2020/2/ 通信不能受信起電力の斥力化
#include <fstream>
#include <iostream>
#include <string>
#include <vector>
#include <stdlib.h>
#include <math.h>
#include <cmath>
#include <sstream>
#include "float.h"

using namespace std;

//重み
#define D 800 //目標位置　初期800
#define R 50  //受信電力　初期50
#define O 0.0002 //通信不能な受信電力座標
//閾値
#define shikiichi_min -249 //受信電力のポテンシャル用、下の閾値、dBm
#define shikiichi_max -60  //上の閾値 dBm
#define RAP 5		   //目標位置のポテンシャルを反転して行う回数

#define X_KINBOU 10 //付近の通信不能座標を検知するときの範囲、現在地+-KINBOU
#define Y_KINBOU 10
#define Z_KINBOU 1
//地図サイズ 使うシミュレーション範囲に応じて変換してください
#define MAPSIZE_X 94
#define MAPSIZE_Y 52
#define MAPSIZE_Z 9

//出発と到着[3] {x,y,z}

//int start[3] = {4,46,0};//
int start[3] = {0,0,0};//
int goal[3] = {59,10,8};//

double received_power[MAPSIZE_X][MAPSIZE_Y][MAPSIZE_Z];
double true_height[MAPSIZE_X][MAPSIZE_Y];

//受信電力値を抜き出したものから座標をつけたtxtをつくる

string file1 = "130.1K2490GD2route_grid.csv"; //グリッド表記経路が書かれた出力ファイルの名前
string file2 = "130.1K249GD2route_meter.csv";//メートル表記経路が書かれた出力ファイルの名前

 vector<string> split(string& input, char delimiter)
	{
    istringstream stream(input);
    string field;
    vector<string> result;
    while (getline(stream, field, delimiter)) {
        result.push_back(field);
    }
    return result;
	}

//受信電力値を抜き出したファイルから座標をつけたファイルをつくる(powerのみ)=>(x,y,z,power)
int input_WI_result(){
	ifstream ifs("d:\\power.txt");//受信起電力が入ったファイル
	ifstream ifs_t("d:\\sabun.csv");//z座標をメートル表記に変換するためのもの
	ifstream ifs_i("d:\\New_Iti.csv");//位置情報ファイル
	//エラー文
	if(!ifs)
	{
		std::cout << "Could not input power.txt" << endl;
		return -1;
	}
	if(!ifs_t)
	{
		std::cout << "Could not input sabun.csv" << endl;
		return -1;
	}
	if(!ifs_i)
	{
		std::cout << "Could not input New_Iti.csv" << endl;
		return -1;
	}

	//読んだファイルを数列に置き換えて出力
	string str;				//文字列クラス
	vector<double> input;	//動的配列クラス

	//グリッドの各受信電力
	while(getline(ifs,str))
	{
		input.push_back(atof(str.c_str()));
	}
	//出力先のファイル
	ofstream fout( "received_power.txt" );
	if(!fout){
		std::cout<<"creative_power.txt could not be open" << endl;
		return -1;
	}
	else std::cout <<"received_power.txt opened." << endl;

	//メートル表記の高さ
	string str_t;
	vector<double> terrain;

	while(getline(ifs_t,str_t)){
		terrain.push_back(atof(str_t.c_str()));
	}
	ofstream fout_t("true_height.csv" );
	if(!fout_t){
		std::cout<<"creative true_height.csv could not be open" << endl;
		return -1;
	}
	else std::cout <<"true_height.csv opened." << endl;
 	
 	//位置情報ファイル
    string line;
    while (getline(ifs_i, line)) 
	{
        
        vector<string> strvec = split(line, ',');
        
    	start[0] = stoi(strvec.at(0));
    	start[1] = stoi(strvec.at(1));
    	start[2] = stoi(strvec.at(2));
    }

	/////////////////

	//received_power[][][]にいれる
	int k;
	int j;
	int i;
	int s = 0;
	//cout << MAPSIZE_X << endl;
	for(k=0 ; k<MAPSIZE_Z ; k++)
	{
		for(j=0 ; j<MAPSIZE_Y ; j++)
		{
			for(i=0 ; i<MAPSIZE_X ; i++)
			{
				if(input[s])
				{
					input[s]=input[s];
					received_power[i][j][k]=input[s];
					s++;
					//cout << received_power[i][j][k]  << endl;
					fout << i << "," << j << "," << k << "," << received_power[i][j][k] << endl;
				}
			}
		}
	}

	//true_height[][]にいれる
	int t = 0;
	int tx;
	int ty;
	for(int j=0 ; j<MAPSIZE_Y ; j++)
	{
		for(int i=0 ; i<MAPSIZE_X ; i++)
		{
			if(terrain[t])
			{
				terrain[t]=terrain[t];
				true_height[i][j] = terrain[t];
				t++;
				//cout << true_height[i][j] << endl;
				tx = i * 5;
				ty = j * 5;
				fout_t << tx << "," << ty << "," << true_height[i][j] << endl;
			}
		}
	}
	fout.close();
	fout_t.close();
	return 0;
}

//ポテンシャルを求める
double get_pot(int x,int y,int z,int loop)
{
	//目標位置の引力ポテンシャル
	double n = (x-goal[0])*(x-goal[0])+(y-goal[1])*(y-goal[1])+(z-goal[2])*(z-goal[2]);
	double pf_target = - loop  / sqrt(n);

	//受信起電力のポテンシャル
	double power = received_power[x][y][z];
	//上の閾値以上のとき
	if(received_power[x][y][z] >= shikiichi_max)
	{
		power = shikiichi_max;
	}

	double pf_received_power = -atan(power - shikiichi_min);

	double d;			//斥力一時保存
	double pf_uncomunication = 0;//この項の斥力の総和

	//案２、現在地から一定の範囲内にある通信不能座標からの斥力のみ考慮
	for(int a = -X_KINBOU; a<= X_KINBOU;a++)
	{
		for(int b = -Y_KINBOU; b<= Y_KINBOU;b++)
		{
			for(int c = -Z_KINBOU; c<= Z_KINBOU;c++)
			{
				//マップ外の座標はループスキップ
				if(x+a < 0 || y+b < 0 || z+c < 0) continue;
				if(x+a >= MAPSIZE_X || y+b >= MAPSIZE_Y || z+c >= MAPSIZE_Z) continue;
				//閾値以下のとき斥力計算
				if(received_power[x+a][y+b][z+c] < shikiichi_min)
				{
					d = sqrt((a*a)+(b*b)+(c*c));
					pf_uncomunication += 1/d;
				}
			}
		}
	}
	//a,b,cループ終了


	//ポテンシャルの和
 	double pot = D * pf_target + R * pf_received_power + O * pf_uncomunication;
	//下の閾値以下、通信不能だと考えられる座標に対して
	if(received_power[x][y][z] <= shikiichi_min)
	{
		pot = 250;
	}
	return pot;
}

//main文
int main(void)
{
	input_WI_result();


	ofstream fout(file1);//グリッドの座標で表される経路
	ofstream fout_t(file2);//実際の値(m表記)で表される経路

	if(!fout)
	{
		cout << file1 << " could not be opened." << endl;
		return -1;
	}
	else cout << file1 << " opened." << endl;

	if(!fout_t)
	{
		cout << file2 << " could not be opened." << endl;
		return -1;
	}
	else cout << file2  << " opened." << endl;
	//出発地点
	int x = start[0];
	int y = start[1];
	int z = start[2];
	int go_z = 10 + (z * 5); //地表高だけ
	//fout << x << "," << y << "," << z << endl;
	//経路
	int max = 100;
	int path_x[100];
	int path_y[100];
	int path_z[100];

	//実際の距離
	int tx = x * 5;
	int ty = y * 5;
	float tz; //地表の起伏+地表高
	fout << tx << "," << ty << "," << go_z << endl;

	int px,py,pz;//次の座標保存
	int loop = 1;//同じ箇所を通った際条件変更用
	int count = 0;//停止用
	int rap = 0;//条件を変えて行う判定
	//int px2,py2,pz2;//第2候補の座標保存　使用していない
	for (int i = 0; i < max; i++)
	{
		if(x!=goal[0] || y!=goal[1] || z!=goal[2])
		{
			double min = DBL_MAX;

			bool p = 0, q = 0, r = 0; //通った座標かのフラグ
			//26近傍から適した経路を求める
			cout << "/ /" << x << " " << y << " " << z << endl ;
			for(int a = -1; a <= 1; a++)
			{
				for(int b = -1; b <= 1; b++)
				{
					for(int c = -1; c <= 1; c++)
					{
						for(int w = 0; w < i; w++)
						{
							if((path_x[w] == x+a && path_y[w] == y+b && path_z[w] == z+c) || (x+a == start[0] && y+b == start[1] && z+c == start[2])) p=1;//path_ []に記録された座標ならp=1
							else p = 0;
						}
						//マップ外の座標はループスキップ
						if(x+a < 0 || y+b < 0 || z+c < 0) continue;
						if(x+a >= MAPSIZE_X || y+b >= MAPSIZE_Y || z+c >= MAPSIZE_Z) continue;
						//20近傍の座標と受信起電力、ポテンシャル
						cout << x+a << " " << y+b << " " << z+c << " " << received_power[x+a][y+b][z+c] << " " << get_pot(x+a,y+b,z+c,loop) << endl;//各近傍の表示

						//最低値を更新したら行う
						if(get_pot(x+a,y+b,z+c,loop) < min)
						{ 
							min = get_pot(x+a,y+b,z+c,loop); // 最低値を新しくする
							px = x+a; //次に経路を考える座標px,py,pz
							py = y+b;
							pz = z+c;
							if(p==1) q = 1;   //p=1　通った座標なら
							else if(min==250) q= 1;
							else q = 0;
						}
					}
				}
			} //a,b,c 処理終了

			if(q==1)
			{
				//loop = -1 ;
				if(rap == 0)
				{
				rap = RAP;//
				}
			}else{
			//loop = 1;
			}

			//loop -1なら目標位置ポテンシャルを反転して同じ箇所で再度行う
			if(rap > 0)
			{		
				loop = -1;
				rap = rap - 1;
                path_x[i] = x;
                path_y[i] = y;
                path_z[i] = z;
                tx = x * 5;
                ty = y * 5;
                tz = (z * 5) + 10 + true_height[x][y];
                go_z = 10 + (z * 5);
				count = count + 1;
				cout << " 次の経路は目標位置のポテンシャルを反転 " << endl;
			}
			else
			{
				loop = 1;
				x = px;
	            y = py;
        	    z = pz;
                path_x[i] = x;
                path_y[i] = y;
                path_z[i] = z;
                tx = x * 5;
                ty = y * 5;
                tz = (z * 5) + 10 + true_height[x][y];
                go_z = 10 + (z * 5);
			}

			if (count == 15)
			{
				i = max;
				printf("ループ終了");
			}

			fout << tx << "," << ty << "," << go_z << "," << received_power[x][y][z] << endl;
			fout_t << tx << "," << ty << "," << tz << "," << received_power[x][y][z] << endl;
		}
	}
	return 0;
}

//�|�e���V�����@��3�����̌o�H�����߂�v���O����
//2020/10/2�@�ア��M�@�d�͂�����Ă����o�H�����߂�
//2020/2/ �ʐM�s�\��M�N�d�͂̐˗͉�
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

//�d��
#define D 800 //�ڕW�ʒu�@����800
#define R 50  //��M�d�́@����50
#define O 0.0002 //�ʐM�s�\�Ȏ�M�d�͍��W
//臒l
#define shikiichi_min -249 //��M�d�͂̃|�e���V�����p�A����臒l�AdBm
#define shikiichi_max -60  //���臒l dBm
#define RAP 5		   //�ڕW�ʒu�̃|�e���V�����𔽓]���čs����

#define X_KINBOU 10 //�t�߂̒ʐM�s�\���W�����m����Ƃ��͈̔́A���ݒn+-KINBOU
#define Y_KINBOU 10
#define Z_KINBOU 1
//�n�}�T�C�Y �g���V�~�����[�V�����͈͂ɉ����ĕϊ����Ă�������
#define MAPSIZE_X 94
#define MAPSIZE_Y 52
#define MAPSIZE_Z 9

//�o���Ɠ���[3] {x,y,z}

//int start[3] = {4,46,0};//
int start[3] = {0,0,0};//
int goal[3] = {59,10,8};//

double received_power[MAPSIZE_X][MAPSIZE_Y][MAPSIZE_Z];
double true_height[MAPSIZE_X][MAPSIZE_Y];

//��M�d�͒l�𔲂��o�������̂�����W������txt������

string file1 = "130.1K2490GD2route_grid.csv"; //�O���b�h�\�L�o�H�������ꂽ�o�̓t�@�C���̖��O
string file2 = "130.1K249GD2route_meter.csv";//���[�g���\�L�o�H�������ꂽ�o�̓t�@�C���̖��O

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

//��M�d�͒l�𔲂��o�����t�@�C��������W�������t�@�C��������(power�̂�)=>(x,y,z,power)
int input_WI_result(){
	ifstream ifs("d:\\power.txt");//��M�N�d�͂��������t�@�C��
	ifstream ifs_t("d:\\sabun.csv");//z���W�����[�g���\�L�ɕϊ����邽�߂̂���
	ifstream ifs_i("d:\\New_Iti.csv");//�ʒu���t�@�C��
	//�G���[��
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

	//�ǂ񂾃t�@�C���𐔗�ɒu�������ďo��
	string str;				//������N���X
	vector<double> input;	//���I�z��N���X

	//�O���b�h�̊e��M�d��
	while(getline(ifs,str))
	{
		input.push_back(atof(str.c_str()));
	}
	//�o�͐�̃t�@�C��
	ofstream fout( "received_power.txt" );
	if(!fout){
		std::cout<<"creative_power.txt could not be open" << endl;
		return -1;
	}
	else std::cout <<"received_power.txt opened." << endl;

	//���[�g���\�L�̍���
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
 	
 	//�ʒu���t�@�C��
    string line;
    while (getline(ifs_i, line)) 
	{
        
        vector<string> strvec = split(line, ',');
        
    	start[0] = stoi(strvec.at(0));
    	start[1] = stoi(strvec.at(1));
    	start[2] = stoi(strvec.at(2));
    }

	/////////////////

	//received_power[][][]�ɂ����
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

	//true_height[][]�ɂ����
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

//�|�e���V���������߂�
double get_pot(int x,int y,int z,int loop)
{
	//�ڕW�ʒu�̈��̓|�e���V����
	double n = (x-goal[0])*(x-goal[0])+(y-goal[1])*(y-goal[1])+(z-goal[2])*(z-goal[2]);
	double pf_target = - loop  / sqrt(n);

	//��M�N�d�͂̃|�e���V����
	double power = received_power[x][y][z];
	//���臒l�ȏ�̂Ƃ�
	if(received_power[x][y][z] >= shikiichi_max)
	{
		power = shikiichi_max;
	}

	double pf_received_power = -atan(power - shikiichi_min);

	double d;			//�˗͈ꎞ�ۑ�
	double pf_uncomunication = 0;//���̍��̐˗͂̑��a

	//�ĂQ�A���ݒn������͈͓̔��ɂ���ʐM�s�\���W����̐˗͂̂ݍl��
	for(int a = -X_KINBOU; a<= X_KINBOU;a++)
	{
		for(int b = -Y_KINBOU; b<= Y_KINBOU;b++)
		{
			for(int c = -Z_KINBOU; c<= Z_KINBOU;c++)
			{
				//�}�b�v�O�̍��W�̓��[�v�X�L�b�v
				if(x+a < 0 || y+b < 0 || z+c < 0) continue;
				if(x+a >= MAPSIZE_X || y+b >= MAPSIZE_Y || z+c >= MAPSIZE_Z) continue;
				//臒l�ȉ��̂Ƃ��˗͌v�Z
				if(received_power[x+a][y+b][z+c] < shikiichi_min)
				{
					d = sqrt((a*a)+(b*b)+(c*c));
					pf_uncomunication += 1/d;
				}
			}
		}
	}
	//a,b,c���[�v�I��


	//�|�e���V�����̘a
 	double pot = D * pf_target + R * pf_received_power + O * pf_uncomunication;
	//����臒l�ȉ��A�ʐM�s�\���ƍl��������W�ɑ΂���
	if(received_power[x][y][z] <= shikiichi_min)
	{
		pot = 250;
	}
	return pot;
}

//main��
int main(void)
{
	input_WI_result();


	ofstream fout(file1);//�O���b�h�̍��W�ŕ\�����o�H
	ofstream fout_t(file2);//���ۂ̒l(m�\�L)�ŕ\�����o�H

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
	//�o���n�_
	int x = start[0];
	int y = start[1];
	int z = start[2];
	int go_z = 10 + (z * 5); //�n�\������
	//fout << x << "," << y << "," << z << endl;
	//�o�H
	int max = 100;
	int path_x[100];
	int path_y[100];
	int path_z[100];

	//���ۂ̋���
	int tx = x * 5;
	int ty = y * 5;
	float tz; //�n�\�̋N��+�n�\��
	fout << tx << "," << ty << "," << go_z << endl;

	int px,py,pz;//���̍��W�ۑ�
	int loop = 1;//�����ӏ���ʂ����ۏ����ύX�p
	int count = 0;//��~�p
	int rap = 0;//������ς��čs������
	//int px2,py2,pz2;//��2���̍��W�ۑ��@�g�p���Ă��Ȃ�
	for (int i = 0; i < max; i++)
	{
		if(x!=goal[0] || y!=goal[1] || z!=goal[2])
		{
			double min = DBL_MAX;

			bool p = 0, q = 0, r = 0; //�ʂ������W���̃t���O
			//26�ߖT����K�����o�H�����߂�
			cout << "/ /" << x << " " << y << " " << z << endl ;
			for(int a = -1; a <= 1; a++)
			{
				for(int b = -1; b <= 1; b++)
				{
					for(int c = -1; c <= 1; c++)
					{
						for(int w = 0; w < i; w++)
						{
							if((path_x[w] == x+a && path_y[w] == y+b && path_z[w] == z+c) || (x+a == start[0] && y+b == start[1] && z+c == start[2])) p=1;//path_ []�ɋL�^���ꂽ���W�Ȃ�p=1
							else p = 0;
						}
						//�}�b�v�O�̍��W�̓��[�v�X�L�b�v
						if(x+a < 0 || y+b < 0 || z+c < 0) continue;
						if(x+a >= MAPSIZE_X || y+b >= MAPSIZE_Y || z+c >= MAPSIZE_Z) continue;
						//20�ߖT�̍��W�Ǝ�M�N�d�́A�|�e���V����
						cout << x+a << " " << y+b << " " << z+c << " " << received_power[x+a][y+b][z+c] << " " << get_pot(x+a,y+b,z+c,loop) << endl;//�e�ߖT�̕\��

						//�Œ�l���X�V������s��
						if(get_pot(x+a,y+b,z+c,loop) < min)
						{ 
							min = get_pot(x+a,y+b,z+c,loop); // �Œ�l��V��������
							px = x+a; //���Ɍo�H���l������Wpx,py,pz
							py = y+b;
							pz = z+c;
							if(p==1) q = 1;   //p=1�@�ʂ������W�Ȃ�
							else if(min==250) q= 1;
							else q = 0;
						}
					}
				}
			} //a,b,c �����I��

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

			//loop -1�Ȃ�ڕW�ʒu�|�e���V�����𔽓]���ē����ӏ��ōēx�s��
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
				cout << " ���̌o�H�͖ڕW�ʒu�̃|�e���V�����𔽓] " << endl;
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
				printf("���[�v�I��");
			}

			fout << tx << "," << ty << "," << go_z << "," << received_power[x][y][z] << endl;
			fout_t << tx << "," << ty << "," << tz << "," << received_power[x][y][z] << endl;
		}
	}
	return 0;
}

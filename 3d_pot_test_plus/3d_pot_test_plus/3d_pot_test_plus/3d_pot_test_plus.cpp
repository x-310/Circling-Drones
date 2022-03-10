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
#define G 0  //���g����Ȃ����̒n��ǁ@����50
#define AA 0  //another drone�@����50
#define O1 0.0002 //�ʐM�s�\�Ȏ�M�d�͍��W
#define O2 0.0002 //�ʐM�s�\�Ȏ�M�d�͍��W
#define O3 0.0002 //�ʐM�s�\�Ȏ�M�d�͍��W
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

double received_power1[MAPSIZE_X][MAPSIZE_Y][MAPSIZE_Z];
double received_power2[MAPSIZE_X][MAPSIZE_Y][MAPSIZE_Z];
double received_power3[MAPSIZE_X][MAPSIZE_Y][MAPSIZE_Z];
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
	ifstream ifs_p1("power1.txt");	//��M�N�d�͂��������t�@�C��
	ifstream ifs_p2("power2.txt");	//��M�N�d�͂��������t�@�C��
	ifstream ifs_p3("power3.txt");	//��M�N�d�͂��������t�@�C��
	ifstream ifs_t("sabun.csv");	//z���W�����[�g���\�L�ɕϊ����邽�߂̂���
	ifstream ifs_i("New_Iti.csv");	//�ʒu���t�@�C��
	//�G���[��
	if (!ifs_p1)
	{
		std::cout << "Could not input power1.txt" << endl;
		return -1;
	}
	if (!ifs_p2)
	{
		std::cout << "Could not input power2.txt" << endl;
		return -1;
	}
	if (!ifs_p3)
	{
		std::cout << "Could not input power3.txt" << endl;
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
	string str;					//������N���X
	vector<double> input_p1;	//���I�z��N���X
	vector<double> input_p2;	//���I�z��N���X
	vector<double> input_p3;	//���I�z��N���X

	//�O���b�h�̊e��M�d��
	while (getline(ifs_p1, str))
	{
		input_p1.push_back(atof(str.c_str()));
	}
	//�o�͐�̃t�@�C��
	ofstream fout_p1( "received_power1.txt" );
	if(!fout_p1){
		std::cout<<"creative_power1.txt could not be open" << endl;
		return -1;
	}
	else std::cout <<"received_power1.txt opened." << endl;

	//�O���b�h�̊e��M�d��
	while (getline(ifs_p2, str))
	{
		input_p2.push_back(atof(str.c_str()));
	}
	//�o�͐�̃t�@�C��
	ofstream fout_p2( "received_power2.txt" );
	if (!fout_p2) {
		std::cout <<"creative_power2.txt could not be open" << endl;
		return -1;
	}
	else std::cout <<"received_power2.txt opened." << endl;

	//�O���b�h�̊e��M�d��
	while (getline(ifs_p3, str))
	{
		input_p3.push_back(atof(str.c_str()));
	}
	//�o�͐�̃t�@�C��
	ofstream fout_p3( "received_power3.txt" );
	if (!fout_p3) {
		std::cout <<"creative_power3.txt could not be open" << endl;
		return -1;
	}
	else std::cout << "received_power3.txt opened." << endl;

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

	//received_power1[][][]�ɂ����
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
				if(input_p1[s])
				{
					input_p1[s]= input_p1[s];
					received_power1[i][j][k]= input_p1[s];
					s++;
					//cout << received_power[i][j][k]  << endl;
					fout_p1 << i << "," << j << "," << k << "," << received_power1[i][j][k] << endl;
				}
			}
		}
	}

	//received_powe2[][][]�ɂ����
	s = 0;
	for (k = 0; k < MAPSIZE_Z; k++)
	{
		for (j = 0; j < MAPSIZE_Y; j++)
		{
			for (i = 0; i < MAPSIZE_X; i++)
			{
				if (input_p2[s])
				{
					input_p2[s] = input_p2[s];
					received_power2[i][j][k] = input_p2[s];
					s++;
					//cout << received_power[i][j][k]  << endl;
					fout_p2 << i << "," << j << "," << k << "," << received_power2[i][j][k] << endl;
				}
			}
		}
	}

	//received_powe3[][][]�ɂ����
	s = 0;
	for (k = 0; k < MAPSIZE_Z; k++)
	{
		for (j = 0; j < MAPSIZE_Y; j++)
		{
			for (i = 0; i < MAPSIZE_X; i++)
			{
				if (input_p3[s])
				{
					input_p3[s] = input_p3[s];
					received_power3[i][j][k] = input_p3[s];
					s++;
					//cout << received_power[i][j][k]  << endl;
					fout_p3 << i << "," << j << "," << k << "," << received_power3[i][j][k] << endl;
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
	fout_p1.close();
	fout_p2.close();
	fout_p3.close();
	fout_t.close();
	return 0;
}

//�|�e���V���������߂�
double get_pot(int x,int y,int z,int loop)
{
	//�ڕW�ʒu�̈��̓|�e���V����
	double n = (x-goal[0])*(x-goal[0])+(y-goal[1])*(y-goal[1])+(z-goal[2])*(z-goal[2]);
	double pf_target = - loop  / sqrt(n);
	////////////////////////////////////////////////////////////////////////
	//��M�N�d�͂̃|�e���V����
	double power1 = received_power1[x][y][z];
	//���臒l�ȏ�̂Ƃ�
	if(received_power1[x][y][z] >= shikiichi_max)
	{
		power1 = shikiichi_max;
	}

	double pf_received_power1 = -atan(power1 - shikiichi_min);

	double d1;			//�˗͈ꎞ�ۑ�
	double pf_uncomunication1 = 0;//���̍��̐˗͂̑��a

	//�ĂQ�A���ݒn������͈͓̔��ɂ���ʐM�s�\���W����̐˗͂̂ݍl��
	for(int a1 = -X_KINBOU; a1<= X_KINBOU;a1++)
	{
		for(int b1 = -Y_KINBOU; b1<= Y_KINBOU;b1++)
		{
			for(int c1 = -Z_KINBOU; c1<= Z_KINBOU;c1++)
			{
				//�}�b�v�O�̍��W�̓��[�v�X�L�b�v
				if(x+a1 < 0 || y+b1 < 0 || z+c1 < 0) continue;
				if(x+a1 >= MAPSIZE_X || y+b1 >= MAPSIZE_Y || z+c1 >= MAPSIZE_Z) continue;
				//臒l�ȉ��̂Ƃ��˗͌v�Z
				if(received_power1[x+a1][y+b1][z+c1] < shikiichi_min)
				{
					d1 = sqrt((a1*a1)+(b1*b1)+(c1*c1));
					pf_uncomunication1 += 1/d1;
				}
			}
		}
	}
	//a,b,c���[�v�I��
	//////////////////////////////////////////////////////////////////////
	//��M�N�d�͂̃|�e���V����
	double power2 = received_power2[x][y][z];
	//���臒l�ȏ�̂Ƃ�
	if (received_power2[x][y][z] >= shikiichi_max)
	{
		power2 = shikiichi_max;
	}

	double pf_received_power2 = -atan(power2 - shikiichi_min);

	double d2;			//�˗͈ꎞ�ۑ�
	double pf_uncomunication2 = 0;//���̍��̐˗͂̑��a

	//�ĂQ�A���ݒn������͈͓̔��ɂ���ʐM�s�\���W����̐˗͂̂ݍl��
	for (int a2 = -X_KINBOU; a2 <= X_KINBOU; a2++)
	{
		for (int b2 = -Y_KINBOU; b2 <= Y_KINBOU; b2++)
		{
			for (int c2 = -Z_KINBOU; c2 <= Z_KINBOU; c2++)
			{
				//�}�b�v�O�̍��W�̓��[�v�X�L�b�v
				if (x + a2 < 0 || y + b2 < 0 || z + c2 < 0) continue;
				if (x + a2 >= MAPSIZE_X || y + b2 >= MAPSIZE_Y || z + c2 >= MAPSIZE_Z) continue;
				//臒l�ȉ��̂Ƃ��˗͌v�Z
				if (received_power2[x + a2][y + b2][z + c2] < shikiichi_min)
				{
					d2 = sqrt((a2 * a2) + (b2 * b2) + (c2 * c2));
					pf_uncomunication2 += 1 / d2;
				}
			}
		}
	}
	//a,b,c���[�v�I��
	//////////////////////////////////////////////////////////////////////////
	////��M�N�d�͂̃|�e���V����
	double power3 = received_power3[x][y][z];
	//���臒l�ȏ�̂Ƃ�
	if (received_power3[x][y][z] >= shikiichi_max)
	{
		power3 = shikiichi_max;
	}

	double pf_received_power3 = -atan(power3 - shikiichi_min);

	double d3;			//�˗͈ꎞ�ۑ�
	double pf_uncomunication3 = 0;//���̍��̐˗͂̑��a

	//�ĂQ�A���ݒn������͈͓̔��ɂ���ʐM�s�\���W����̐˗͂̂ݍl��
	for (int a3 = -X_KINBOU; a3 <= X_KINBOU; a3++)
	{
		for (int b3 = -Y_KINBOU; b3 <= Y_KINBOU; b3++)
		{
			for (int c3 = -Z_KINBOU; c3 <= Z_KINBOU; c3++)
			{
				//�}�b�v�O�̍��W�̓��[�v�X�L�b�v
				if (x + a3 < 0 || y + b3 < 0 || z + c3 < 0) continue;
				if (x + a3 >= MAPSIZE_X || y + b3 >= MAPSIZE_Y || z + c3 >= MAPSIZE_Z) continue;
				//臒l�ȉ��̂Ƃ��˗͌v�Z
				if (received_power3[x + a3][y + b3][z + c3] < shikiichi_min)
				{
					d3 = sqrt((a3 * a3) + (b3 * b3) + (c3 * c3));
					pf_uncomunication3 += 1 / d3;
				}
			}
		}
	}
	//a,b,c���[�v�I��
	////////////////////////////////////////////////////////////////////////

	//�|�e���V�����̘a
	if (isinf(pf_uncomunication3)) {
		pf_uncomunication3 = 1;
	}
	double pot = D * pf_target + R * pf_received_power1 + O1 * pf_uncomunication1 - G * pf_received_power2 - O2 * pf_uncomunication2 - AA * pf_received_power3 - O3 * pf_uncomunication3;
	//����臒l�ȉ��A�ʐM�s�\���ƍl��������W�ɑ΂���
	if(received_power1[x][y][z] <= shikiichi_min)
	{
		pot = 250;
	}
	return pot;
	if (received_power2[x][y][z] <= shikiichi_min)
	{
		pot = 250;
	}
	return pot;
	if (received_power3[x][y][z] <= shikiichi_min)
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
						cout << x+a << " " << y+b << " " << z+c << " " << received_power1[x+a][y+b][z+c] << " " << get_pot(x+a,y+b,z+c,loop) << endl;//�e�ߖT�̕\��

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

			fout << tx << "," << ty << "," << go_z << "," << received_power1[x][y][z] << endl;
			fout_t << tx << "," << ty << "," << tz << "," << received_power1[x][y][z] << endl;
		}
	}
	return 0;
}

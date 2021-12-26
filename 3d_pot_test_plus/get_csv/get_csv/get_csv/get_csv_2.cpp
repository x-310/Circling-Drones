#include <fstream>
#include <iostream>
#include <string>
#include <vector>
#include <stdlib.h>
#include <math.h>
#include <cmath>
#include <sstream>
#include "float.h"

int start[3];//

using namespace std;

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

int input_WI_result() {
	ifstream ifs_i("d:\\New_Iti.csv");//位置情報ファイル
	//エラー文
	if (!ifs_i) {
		cout << "Could not input New_Iti.csv" << endl;
		return -1;
	}
	cout << "New_Iti.csv open" << endl;

	//位置情報ファイル
	string str_i;
	vector<int> iti;
	//int iCnt = 0;
	while (getline(ifs_i, str_i)) {
		iti.push_back(atoi(str_i.c_str()));

		vector<string> strvec = split(str_i, ',');

		for (int i = 0; i < strvec.size(); i++) {
			//printf("%5d\n", stoi(strvec.at(i)));
			start[i] = stoi(strvec.at(i));
		}
	}
	//start[0] = 4;
	//start[1] = 46;
	//start[2] = 0;
}

int main()
{
	input_WI_result();

	int x = start[0];
	int y = start[1];
	int z = start[2];
}
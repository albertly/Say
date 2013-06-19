// Say++.cpp : Defines the entry point for the console application.

// Copyright 2007-2009 Martin Krolik
/*
    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with this program.  If not, see <http://www.gnu.org/licenses/>.
*/



#include "stdafx.h"
//#include <atlbase.h>
#include <sapi.h>           // SAPI includes
#include <sphelper.h>
//using namespace sapi;

#include <iostream>
#include <fstream>
#include <string>

#include <locale>
//#include <codecvt>
#include <cstdlib>
#include <time.h> 
using namespace std;

wstring array[5];
wstring rate[5];
wstring weight[5];
int size_a = 0;
int i_rate[5] = {0,0,0,0,0};
int i_weight[5] = {0,0,0,0,0};

int dice()
{
  /* initialize random seed: */
  srand (time(NULL));
  
  return rand() % 100;
}

int rollDice()
{
	int i_dice = dice();
	int sum = 0;
	for (int i=0; i< size_a; i++)
	{
		sum += i_weight[i];
		if (sum >= i_dice)
		{
			return i;
		}
	}
}
void writeToLog(char * st)
{
	return;
	ofstream myfile;
	myfile.open ("error_log.txt", ios::app | ios::out);
	myfile << st;
	myfile.close();
}

void wstr2int(wstring a[], int i_a[])
{
	for (int i=0; i< size_a; i++)
		i_a[i] = _wtoi(a[i].c_str());
}

int readFileToArray(char* f, wstring a[]) 
{	
	
	//const std::locale empty_locale = std::locale::empty();
    //typedef std::codecvt_utf8<wchar_t> converter_type;
    //const converter_type* converter = new converter_type;
	
	int loop=0; //Index of array
	wstring line; //The line taken from the *.txt source
	wifstream  myfile (f); //To read from the *.txt File
//	stream.imbue(utf8_locale);
	
	if (myfile.is_open()) //Checking if the file can be opened
	{
		while (! myfile.eof() ) //Runs while the file is NOT at the end
		{
			getline (myfile,line); //Gets a single line from example.txt
			a[loop]=line; //Saves that line in the array
			loop++; //Does an increment to the variable 'loop'
		}
		myfile.close(); //Closes the file
		return loop;
	}
	else 
		writeToLog("Unable to open file\n");
		return -1;

}
int _tmain(int argc, _TCHAR* argv[])
{
	wstring wstr = L"-20";
	int number = _wtoi(wstr.c_str());
	ofstream myfile;
	int index = 0;

	unsigned long ulngVoiceNumber = 0;
	ISpVoice * spvoice;

	//HRESULT				hr = S_OK;
	CComPtr <ISpVoice>		cpVoice;
	CComPtr <ISpObjectToken>	cpToken;
	CComPtr <IEnumSpObjectTokens>	cpEnum;

	int iSize = 0;
	for (int i = 0; i < argc; i++)
	{
		iSize = iSize + _tcslen(argv[i]) + 1;
	}
	writeToLog("params \n");

	size_a = readFileToArray(".\\voices.ini", array);
	readFileToArray(".\\rate.ini", rate);
	wstr2int(rate, i_rate);
	readFileToArray(".\\weight.ini", weight);
	wstr2int(weight, i_weight);
	
	index = rollDice();

	writeToLog("read file to array \n");


	TCHAR * wstrText ; 

	wstrText = new TCHAR[iSize];

	ZeroMemory(wstrText, sizeof(wstrText));

	for (int i = 1; i < argc; i++)
	{
		_tcscat(wstrText, argv[i]);
		_tcscat(wstrText, L" ");
	}
	

	HRESULT hr = CoInitialize(NULL); 
	
	CLSID clsid ;
	hr = CLSIDFromProgID(OLESTR("SAPI.SpVoice"), &clsid);

	//Create a SAPI voice
	//hr = cpVoice.CoCreateInstance( CLSID_SpVoice );
	if(SUCCEEDED(hr))
	{
		hr = CoCreateInstance(clsid, NULL, CLSCTX_INPROC_SERVER, __uuidof(ISpVoice), (LPVOID *)&cpVoice);
	}
	else
	{		
		writeToLog("CoCreateInstance\n");
	}

	//Enumerate voice tokens with attribute "Name=Microsoft Sam” 
	if(SUCCEEDED(hr))
	{
		//IVONA Jennifer
		//Name=Heather22k
		wstring stVoice = array[index]; // L"Name=Heather22k";
		
		//hr = SpEnumTokens(L"HKEY_CURRENT_USER\\Software\\Microsoft\\Speech\Voices", L"Name=Heather22k", NULL, &cpEnum);
		hr = SpEnumTokens(SPCAT_VOICES, stVoice.data(), NULL, &cpEnum);
		//hr = SpEnumTokens(SPCAT_VOICES, L"Name=Microsoft Sam", NULL, &cpEnum);
	}
	else
	{
		writeToLog("SpEnumTokens.\n");
	}
    
	//Get the closest token
	if(SUCCEEDED(hr))
	{
		hr = cpEnum ->Next(1, &cpToken, NULL);
	}
	else
	{
		writeToLog("Next.\n");
	}

	//set the voice 
	if(SUCCEEDED(hr))
	{
		hr = cpVoice->SetVoice( cpToken);
		hr = cpVoice->SetRate((long)i_rate[index]); 
	}
	else
	{
		writeToLog("SetVoice.\n");
	}

	//set the output to the default audio device
	if(SUCCEEDED(hr))
	{
		hr = cpVoice->SetOutput( NULL, TRUE );
	}
	else
	{
		writeToLog("SetOutput.\n");
	}
		
	//Speak the text file (assumed to exist)
	if(SUCCEEDED(hr))
	{
		cpVoice->Speak(wstrText, SpeechVoiceSpeakFlags::SVSFDefault, &ulngVoiceNumber);
	}
	else
	{
		writeToLog("Speak.\n");
	}

	//Release objects
	cpVoice.Release ();
	cpEnum.Release();
	cpToken.Release();
	
	CoUninitialize();

	delete [] wstrText;

	writeToLog("End.\n");

	return 0;

	///////////////////////////////////////////////////////////////////

/////////////////////////////////////////////////////////	
//	HRESULT hr = CoInitialize(NULL); 
//	
//	CLSID clsid ;
//	hr = CLSIDFromProgID(OLESTR("SAPI.SpVoice"), &clsid);
/////////////////////////////////////////////////////////////////////////////
	if (FAILED(hr))
	{
		printf("Failed to retrive CLSID for COM server");
		return -1;
	}

	

	hr = CoCreateInstance(clsid, NULL, CLSCTX_INPROC_SERVER, __uuidof(ISpVoice), (LPVOID *)&spvoice);
	
	if (FAILED(hr))
	{
		printf("Failed to start COM server");
		return -1;
	}
	

	spvoice->Speak(wstrText, SpeechVoiceSpeakFlags::SVSFDefault, &ulngVoiceNumber);

	CoUninitialize();

	delete [] wstrText;
	
	return 0;
}




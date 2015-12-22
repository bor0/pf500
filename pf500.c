/*
This file is part of PF500.

PF500 is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

PF500 is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with PF500. If not, see <http://www.gnu.org/licenses/>.

*/

////////////////////////////////////////
//
// COM communication by Boro Sitnikovski
// Protocol for Synergy PF500
//
// Revision: 24.12.2011
// Revision: 27.12.2011
//

#include <windows.h>
#include <stdio.h>

HANDLE hCom;
FILE *inFile, *outFile;

#pragma comment(lib, "user32.lib")

char *initCommand(char command, char *commandData) {
	static char seq = ' ';
	unsigned char data_prefix[] = "\x01%%c%c%c%s\x05%%c%%c%%c%%c\x03";
	unsigned static char data[255];
	unsigned char data2[255];
	unsigned int checkSum=0, len, i, bytwrite, bytread, tries;
	unsigned char bcc[4];

	sprintf(data, data_prefix, seq, command, commandData);

	for (len=1;data[len] != 0x05;len++) checkSum += data[len];

	len += 32 - 1;
	checkSum = checkSum - '%' - 'c' + len + 0x05;

	i = (checkSum&0xF) + 0x30; bcc[0] = i;
	i = ((checkSum&0xF0) >> 4) + 0x30; bcc[1] = i;
	i = ((checkSum&0xF00) >> 8) + 0x30; bcc[2] = i;
	i = ((checkSum&0xF000) >> 12) + 0x30; bcc[3] = i;

	sprintf(data2, data, len, bcc[3], bcc[2], bcc[1], bcc[0]);

	if (seq == '#') seq = ' ';
	else seq = '#';

	len = strlen(data2);

	i = 0;
	while (i < len) {
		WriteFile(hCom, data2 + i, 1, &bytwrite, NULL);
		if (bytwrite == 1) i++;
	}
	
	Sleep(60);

	data[0] = '\0';

	for (tries = 0; tries < 5; tries++) {
		if (ReadFile(hCom, data, 1, &bytread, NULL)) {
			if (data[0] == 0x16) {
				tries--;
				Sleep(60);
			}
			else if (data[0] == 0x01) {
				len = 1;
				while (data[len-1] != 0x03) ReadFile(hCom, data+(len++), 1, &bytread, NULL);
				data[len] = '\0';
				break;
			}
		}
	}
	
	if (data[0] == '\0') strcpy(data + 4, "ERROR: No bytes to read");
	fprintf(outFile, "%s\n", data + 4);

	return data;
}

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nShowCmd) {

	DCB dcb;
	COMMTIMEOUTS cto;
	BOOL fSuccess;
	char *pcCommPort, *arg2, *arg3;
	char line[512]; int i;

	"SynergyPF500 Communicator by Boro Sitnikovski\n-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-";

	pcCommPort = strtok(lpCmdLine, " ");
	arg2 = strtok(NULL, " ");
	arg3 = strtok(NULL, " ");

	if (pcCommPort == NULL || arg2 == NULL || arg3 == NULL) return 0;

	inFile = fopen(arg2, "rb");
	if (!inFile) return 0;

	outFile = fopen(arg3, "w+");
	if (!outFile) {
		fclose(inFile);
		return 1;
	}

	hCom = CreateFile(pcCommPort,
	GENERIC_READ | GENERIC_WRITE,
	FILE_SHARE_READ | FILE_SHARE_WRITE, // must be opened with exclusive-access
	NULL, // no security attributes
	OPEN_EXISTING, // must use OPEN_EXISTING
	0, // not overlapped I/O
	NULL // hTemplate must be NULL for comm devices
	);

	if (hCom == INVALID_HANDLE_VALUE) {
		// Handle the error.
		fprintf(outFile, "ERROR: CreateFile failed (%d)\n", GetLastError());
		return 2;
	}

	// Build on the current configuration, and skip setting the size
	// of the input and output buffers with SetupComm.

	fSuccess = GetCommState(hCom, &dcb);

	if (!fSuccess) {
		// Handle the error.
		fprintf(outFile, "ERROR: GetCommState failed (%d)\n", GetLastError());
		return 3;
	}

	// Fill in DCB: 9,600 bps, 8 data bits, no parity, and 1 stop bit.

	dcb.BaudRate = CBR_9600; // set the baud rate
	dcb.ByteSize = 8; // data size, xmit, and rcv
	dcb.Parity = NOPARITY; // no parity bit
	dcb.StopBits = ONESTOPBIT; // one stop bit

	fSuccess = SetCommState(hCom, &dcb);

	if (!fSuccess) {
		// Handle the error.
		fprintf(outFile, "ERROR: SetCommState failed (%d)\n", GetLastError());
		return 4;
	}

	fSuccess = GetCommTimeouts(hCom, &cto);

	if (!fSuccess) {
		// Handle the error.
		fprintf(outFile, "ERROR: GetCommTimeouts failed (%d)\n", GetLastError());
		return 5;
	}

	cto.ReadTotalTimeoutConstant = 500; // 500 ms

	fSuccess = SetCommTimeouts(hCom, &cto);

	if (!fSuccess) {
		// Handle the error.
		fprintf(outFile, "ERROR: SetCommTimeouts failed with error (%d)\n", GetLastError());
		return 6;
	}

	while (!feof(inFile)) {
		char dataa[511212];
		fgets(line, 200, inFile);
		for (i=0;line[i]!='\r' && line[i] != '\n' && line[i] != '\0';i++);
		line[i] = '\0';
		if (strlen(line) == 0) continue;
		initCommand(line[0], line+1);
		line[0] = '\0';
	}

	fclose(inFile);
	fclose(outFile);
	CloseHandle(hCom);

	return 0;
}

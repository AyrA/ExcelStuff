#include <stdlib.h>
#include <stdio.h>

int printFile(const char* fName);

int main(int argc,char **argv)
{
	if(argc>2)
	{
		switch(decryptFile(argv[1],argv[2]))
		{
			case -1:
				printf("Error opening File %s\r\n",argv[1]);
				break;
			case -2:
				printf("Error opening File %s\r\n",argv[2]);
				break;
			default:
				printf("Password removed\r\nDo the folowing now:\r\n");
				printf("1. Open the excel sheet and press [ALT]+[F11]\r\n");
				printf("   Confirm any errors that might appear\r\n");
				printf("2. go to Tools > VBA project properties.\r\n");
				printf("3. In the Tab \"Protection\" enter any password\r\n");
				printf("   Do not clear the Checkbox!\r\n");
				printf("4. Save, close the Editor, save the excel sheet and close it\r\n");
				printf("5. Open it again\r\n");
				printf("6. Repear step 1 and 2, there should be no errors\r\n");
				printf("7. Clear the password checkbox in the \"protection\" tab\r\n");
				printf("8. Repeat Steps 4 and 5.\r\n");
				printf("9. Your password is gone!\r\n");
				printf("\r\n");
				printf("/u/AyrA_ch\r\n");
				printf("\r\n");
				break;
		}
	}
	else
	{
		printf("excelDecrypt <input-filename> <output-filename>");
	}
	printf("\r\n");
	return 0;
}

int decryptFile(const char* fName1,const char* fName2)
{
	char* cc;
	FILE* fp;
	int fs,c,i;
	if(fp=fopen(fName1,"rb"))
	{
		fseek(fp, 0L, SEEK_END);
		cc=malloc(fs=ftell(fp));
		fseek(fp, 0L, SEEK_SET);
		fread(cc,sizeof(cc[0]),fs,fp);
		for(i=0;i<fs-4;i++)
		{
			if(cc[i]=='D' && cc[i+1]=='P' && cc[i+2]=='B' && cc[i+3]=='=' && cc[i+4]=='"')
			{
				cc[i+2]='x';
			}
		}
		fclose(fp);
		
		if(fp=fopen(fName2,"wb"))
		{
			fwrite(cc,sizeof(cc[0]),fs,fp);
			fclose(fp);
		}
		else
		{
			free(cc);
			return -2;
		}
		free(cc);
	}
	else
	{
		return -1;
	}
	return 0;
}

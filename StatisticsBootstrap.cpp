/**************************************************************************************************
Purpose: Performs a bootstrap resampling techniqhues to estimate the sampling distribution using 
bootstrap. When "Mean" is selected as the sampling statistic, jackknife resampling is used to 
correct standard deviation bias. Also includes a permutation test for mean comparison of 2 samples.

Auther: Ben Stadnick
**************************************************************************************************/
 
#include <Origin.h>


typedef double (*TESTSTATISTICFUN)(vector<double>); // argument passing functions

//Input for bootstrap function using column input, that is sample data is organized into columns
//e.g. Sample 1 data is in column 0, Sample 2 data is in column 1 and Sample 3 data is in column 2.
//All data must be within a single worksheet indicated by "WorkBookName" and "TestDataSheetIndex"
//The bootstrap data will output to a new worksheet. Plotting this data as a histogram will give
//an estimate of the sampling distribution.
void BootstrapInput(){
	vector<string> BootOptions(1);
	BootOptions[0]= "Mean";//Test statistic, available options inlcude "Mean" and "Median"
	
	string WorkBookName = "Book1";//Name of workbook where data is stored
	int TestDataSheetIndex = 0;//index of sheet where data is stored
	int ColumnStartIndex = 0;//index of first column to bootstrap
	int ColumnEndIndex = 0;//index of last column to bootstrap
	int NumberOfResamples = 5249;//Number of times the data will be resampled for the bootstrap
	
	BootstrapColumnInput( BootOptions, WorkBookName, TestDataSheetIndex, ColumnStartIndex, ColumnEndIndex, NumberOfResamples)
}

//Input from row vector
void JackBootMeanRowInput(){
	//User entered control variables
	
	//Select workbook
	string WorkSheetName = "Book1";
	int TestDataSheetIndex = 0; //sheet where data is stored
	int ColumnStartIndex = 0;
	int ColumnEndIndex = 1;
	int CurrentRowIndex = 24;
	
	
	vector<double> SampleData(ColumnEndIndex-ColumnStartIndex);
	WorksheetPage wksPage(WorkSheetName);
	Worksheet wks = wksPage.Layers(TestDataSheetIndex);
	
	for(int kk=0; kk<ColumnEndIndex-ColumnStartIndex; kk++)
	{
		Column X_Col(wks, kk+ColumnStartIndex);
		vector<double> X_Vec = X_Col.GetDataObject();
		SampleData[kk] = X_Vec[CurrentRowIndex];
	}
	//Bootstrap
	vector<double> Means;
	Means = JackBootMean(SampleData, 50000);
	
	
	//Add data to new sheet
	int DataSumIndex = wksPage.AddLayer("Bootstrap");
	Worksheet DataSumSheet = wksPage.Layers(DataSumIndex);
	
	//Mean Values values
		DataSumSheet.Columns(1).SetLongName("Bootstrap"); 
		Dataset<double> ds2(DataSumSheet, (1)); 
		ds2 = Means;
}

//input from column vector
void BootstrapColumnInput(vector<string> BootOptioins, string WorkBookName, int TestDataSheetIndex, int ColumnStartIndex, int ColumnEndIndex, int NumberOfResamples){
	DWORD StartTime = GetTickCount();//For timing
	
	//Select workbook and sheet
	WorksheetPage wksPage(WorkBookName);
	Worksheet wks = wksPage.Layers(TestDataSheetIndex);
	
	//Add data to new sheet to store bootstrap
	int DataSumIndex = wksPage.AddLayer(wks.GetName() + " Bootstrap");
	Worksheet DataSumSheet = wksPage.Layers(DataSumIndex);
	DataSumSheet.Columns(0).SetType(OKDATAOBJ_DESIGNATION_Y);
	
	//initialize vector to hold sample data and resampled means
	vector<double> SampleData(wks.GetNumRows()), ResampledTestStatistics(NumberOfResamples);
	int CurrentColumnIndex;
	
	//
	vector<string> BootStringOtions;
	vector<int> BootIntOptions;
	
	//Loop through columns and bootstrap
	for(CurrentColumnIndex=ColumnStartIndex; CurrentColumnIndex<=ColumnEndIndex; CurrentColumnIndex++)
	{
		//Get current sample data
		Column Data_Col(wks, CurrentColumnIndex);
		vector<double> SampleData = Data_Col.GetDataObject();

		//Perform bootstrap
		if(BootOptioins[0] == "Mean" || BootOptioins[0] == "mean")
			ResampledTestStatistics = JackBootMean(SampleData, NumberOfResamples);
		
		if(BootOptioins[0] == "Median" || BootOptioins[0] == "median")
			ResampledTestStatistics = Bootstrap(SampleData, NumberOfResamples, GetMedianVal, BootStringOtions,  BootIntOptions);
		
		ResampledTestStatistics.Sort();
		
		//Add resampled mean values to new worksheet
		if(CurrentColumnIndex>1)
			DataSumSheet.AddCol();
	
		DataSumSheet.Columns(CurrentColumnIndex-ColumnStartIndex).SetLongName(Data_Col.GetLongName());
		DataSumSheet.Columns(CurrentColumnIndex-ColumnStartIndex).SetUnits(Data_Col.GetUnits()); 
		DataSumSheet.Columns(CurrentColumnIndex-ColumnStartIndex).SetComments(Data_Col.GetComments()); 
		Dataset<double> ds(DataSumSheet, CurrentColumnIndex-ColumnStartIndex); 
		ds = ResampledTestStatistics;
	}
	printf("Processing time was %d ms\n",(GetTickCount()-StartTime));
}

// Bootstrap with jacknife correction; Optimized for mean calculation; 
// Jackknife correction removes events from the data pool to correct for bias in the bootstrap
vector<double> JackBootMean(vector<double> SamplesVec, int NumberOfResamples){
	const int SampleVecLength = SamplesVec.GetSize();
	int ResampleCycles = NumberOfResamples/SampleVecLength;
	
	int nn, kk;//indices
	double Mean = 0;
	int RandNumber;
	vector<double> MeanVec(ResampleCycles*SampleVecLength+SampleVecLength);
	vector<double> JackSampleVec();
	
	//Generate bootstrapped means
	for(kk=0; kk<=ResampleCycles; kk++)
	{
		for(int jj=0; jj<SampleVecLength; jj++){
			Mean = 0;
			JackSampleVec = SamplesVec;
			JackSampleVec.RemoveAt(jj);
			for(nn=0; nn<SampleVecLength; nn++){
				Mean = Mean + JackSampleVec[rnd()*(SampleVecLength-1)];// value is tuncated to an int
			}
		MeanVec[kk*SampleVecLength + jj] = Mean/(SampleVecLength);//Add mean of current resample to list
		}
	}	
return MeanVec;
}

// Bootstrap function for arbitrary test statistic given by TestStatCalc
vector<double> Bootstrap(vector<double> SamplesVec, int NumberOfResamples, TESTSTATISTICFUN TestStatCalc, vector<string> BootStringOtions, vector<int> BootIntOtions){
	//declare variables
	int SampleVecLength = SamplesVec.GetSize(), kk, nn;
	double TestStatisticVal;
	vector<double> TestStatisticVec(NumberOfResamples), ResampleVec(SampleVecLength);
	
	for(kk=0; kk<NumberOfResamples; kk++){
		//resample data
		for(nn=0; nn<SampleVecLength; nn++){
			ResampleVec[nn] = SamplesVec[rnd()*(SampleVecLength)];//"rnd()*(SampleVecLength)" value is tuncated to an int
		}
		
		TestStatisticVal = TestStatCalc(ResampleVec);//calculate test statistic for current resample
		TestStatisticVec[kk] = TestStatisticVal;
	}	
return TestStatisticVec;
}


//input from column vectors
void PermutationColumnInput(){
	DWORD StartTime = GetTickCount();//For timing
	
	int TestDataSheetIndex = 1; //sheet where data is stored
	
	//Select workbook and sheet
	string WorkSheetName = "Book1";
	WorksheetPage wksPage(WorkSheetName);
	Worksheet wks = wksPage.Layers(TestDataSheetIndex);
	
	//Add data to new sheet to store bootstrap
	int DataSumIndex = wksPage.AddLayer(wks.GetName() + " Permutation");
	Worksheet DataSumSheet = wksPage.Layers(DataSumIndex);
	DataSumSheet.Columns(0).SetType(OKDATAOBJ_DESIGNATION_Y);
	
	//Initialize index values
	int ColumnIndex1 = 14;
	int ColumnIndex2 = 15;
	int NumberOfResamples = 999999;
	
	//initialize vector to hold sample data
	vector<double> SampleData(wks.GetNumRows());
	

		Column Data_Col1(wks, ColumnIndex1);
		vector<double> Sample1Data = Data_Col1.GetDataObject();
		
		Column Data_Col2(wks, ColumnIndex2);
		vector<double> Sample2Data = Data_Col2.GetDataObject();
		
		//Perform permutation
		vector<double> PermResults(NumberOfResamples);
		PermResults = Permutation(Sample1Data, Sample2Data, NumberOfResamples);
		PermResults.Sort();
		
		//Mean Values values
		DataSumSheet.Columns(0).SetLongName(Data_Col1.GetName()); 
		Dataset<double> ds2(DataSumSheet, 0); 
		ds2 = PermResults;
	
	
	printf("Processing time was %d ms\n",(GetTickCount()-StartTime));
	
}

// Permutation function
vector<double> Permutation(vector<double> SamplesVec1, vector<double> SamplesVec2,  int NumberOfResamples){
	int SampleVec1Length = SamplesVec1.GetSize();
	int SampleVec2Length = SamplesVec2.GetSize();
	int PooledResultsLength = SampleVec1Length+SampleVec2Length;
	int RandSampleIndex;
	vector<double> PooledResults(PooledResultsLength), PermVec1(SampleVec1Length), PermVec2(SampleVec2Length), DiffVec(NumberOfResamples);
	
	//Pool data into one vector
	for(int ii=0; ii<SampleVec1Length; ii++)
	{
		PooledResults[ii] = SamplesVec1[ii];
	}
	for(ii; ii<PooledResultsLength; ii++)
	{
		PooledResults[ii] = SamplesVec2[ii-SampleVec1Length];
	}
	
	//Resample data
	for(int kk=0; kk<NumberOfResamples; kk++)
	{
		for(int nn=0; nn<SampleVec1Length; nn++)
		{
			RandSampleIndex = rnd()*(PooledResultsLength);//value is tuncated to an int
			PermVec1[nn] = PooledResults[RandSampleIndex]
		}
		for(nn=0; nn<SampleVec2Length; nn++)
		{
			RandSampleIndex = rnd()*(PooledResultsLength);//value is tuncated to an int
			PermVec2[nn] = PooledResults[RandSampleIndex]
		}
		DiffVec[kk] = GetMeanDifference(PermVec1, PermVec2);
	}	
return DiffVec;
}

double GetMeanDifference(vector<double> SamplesVec1, vector<double> SamplesVec2){
	double Mean1=0;
	double Mean2=0;
	
	//Get the mean value of vector 1
	for(int kk=0; kk<SamplesVec1.GetSize(); kk++)
	{
		Mean1 = Mean1 + SamplesVec1[kk];
	}
	Mean1 = Mean1/kk;
	
	//Get the mean value of vector 2
	for(kk=0; kk<SamplesVec2.GetSize(); kk++)
	{
		Mean2 = Mean2 + SamplesVec2[kk];
	}
	Mean2 = Mean2/kk;
	
	return Mean2-Mean1;
}

double GetMedianVal(vector<double> MedianVec){//Note: this function sorts the data and is horribly inefficient
	double Median;
	MedianVec.Sort();
	
	if(mod(MedianVec.GetSize(), 2) == 0)//if even
	Median = (MedianVec[MedianVec.GetSize()/2] + MedianVec[MedianVec.GetSize()/2+1])/2;
	
	if(mod(MedianVec.GetSize(), 2) == 1)//if odd
	Median = MedianVec[MedianVec.GetSize()/2]+1;
	
	return Median;	
}
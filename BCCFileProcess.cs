using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Diagnostics;
using System.Data;
namespace BCCTest
{
    public class BCCFileProcess
    {
        public static void Run()
        {
            ExecutePocess("Job4", @"C:\HigmarkAutomation2014_Jobs\", 63);
        }

        #region "    BCC Process...    "

        static bool ExecutePocess(string strjobNbr, string strjobsfldr, int ircdCnt)
        {
            try
            {

                Console.WriteLine(string.Format("Process started."));
                //////////////////////////////////////////////////////////////////////////////////////////
                ///////////////////////// CREATING THE JOB FOLDER AND ALL NEEDED SUBFOLDERS FOR PROCESSING
                //////////////////////////////////////////////////////////////////////////////////////////
                string strJobNumber = "";// <I BELIEVE YOU'LL BE USING THE ORDER NUMBER FOR THIS - STEVE TO CONFIRM>;
                string strJobsFolder = "";// "\\server\jobs\"  <I BELIEVE THIS WILL BE CHANGING FROM THIS VALUE BECAUSE OF THE NEW FOLDER STRUCTURE - STEVE TO CONFIRM>;

                strJobNumber = strjobNbr;
                strJobsFolder = strjobsfldr;

                string strRootJobFolder = strJobsFolder + strJobNumber + "\\";

                // Verifying that the job folder does not already exist.
                if (Directory.Exists(strRootJobFolder))
                {
                    //LogFile("A job folder for " + strJobNumber + " already exists.", true);
                    //return false;
                }

                // Subfolders of the root job folder.
                string strCSRJobFolder = strRootJobFolder + "CSR\\";
                string strDataPrepJobFolder = strRootJobFolder + "DataPrep\\";
                string strPNetImagesJobFolder = strRootJobFolder + "PNetImages\\";
                string strPrepressJobFolder = strRootJobFolder + "Prepress\\";
                string strProductionJobFolder = strRootJobFolder + "Production\\";

                // Subfolders of the DataPrep folder.
                string strDataJobFolder = strDataPrepJobFolder + "Data\\";
                string strDocsJobFolder = strDataPrepJobFolder + "Docs\\";
                string strProgramsJobFolder = strDataPrepJobFolder + "Programs\\";

                // Subfolders of the Data folder.
                string strBCCJobFolder = strDataJobFolder + "BCC\\";
                string strInputJobFolder = strDataJobFolder + "Input\\";
                string strPresortDataJobFolder = strDataJobFolder + "PresortData\\";
                string strWorkingJobFolder = strDataJobFolder + "Working\\";

                Console.WriteLine(string.Format("before creating directories."));

                // Creating the specified folder/subfolders.
                Directory.CreateDirectory(strCSRJobFolder);
                Directory.CreateDirectory(strBCCJobFolder);
               // Directory.CreateDirectory(strInputJobFolder);
                Directory.CreateDirectory(strPresortDataJobFolder);
                Directory.CreateDirectory(strWorkingJobFolder);
                Directory.CreateDirectory(strDocsJobFolder);
                Directory.CreateDirectory(strProgramsJobFolder);
                Directory.CreateDirectory(strPNetImagesJobFolder);
                Directory.CreateDirectory(strPrepressJobFolder);
                Directory.CreateDirectory(strProductionJobFolder);

                Console.WriteLine(string.Format("Created directories."));

                //////////////////////////////////////////////////////////////////////////////////////////
                /////////////////////////////////////////////////////////////////// INITIALIZING VARIABLES
                //////////////////////////////////////////////////////////////////////////////////////////

                ////
                //string strPresortDataJobFolder = string.Empty;
                //string strBCCJobFolder = string.Empty;
                string strInvalidsRemoved = string.Empty;
                //string strWorkingJobFolder = string.Empty;
                //string strDocsJobFolder = string.Empty;
                int iInputRecords = ircdCnt;
                ////

                // Static values - please define these in a config file.
                string strMailDatContactName = "John Kubiak";
                string strMailDatContactEmail = "mailing@heeter.com";
                string strMailDatContactPhone = "7247468900";
                string strHeeterMailerID = "899477";
                string strMailDatVersion = "12-2";
                string strImportName = "Highmark Fulfillment Import";
                string strNCOACompany = "Highmark Inc.";
                string strPresortSettings = "Highmark Fulfillment Presort";
                //string strOutputLabel = "NO TRACKING - Heeter Standard Output";     
                string strTaskmasterJobsFolder = @"\\gmc-server\bcc\fs\jobs\";
                string strMailDatSettingsFolder = @"\\gmc-server\bcc\fs\settings\maildat\";
                string strProjectName = "Highmark Fulfillment";
                string strBCCMailManEXE = @"\\gmc-server\bcc\fs\MailManStub.exe";
                string strOutputLabel = "NO TRACKING - Heeter Standard Output";

                // string strJobNumber = ""; //<I BELIEVE YOU'LL BE USING THE ORDER NUMBER FOR THIS - STEVE TO CONFIRM>;
                string strInputDataName = Path.Combine(strInputJobFolder, "Highmark_Sample_Data_Feed_BCCInput.txt"); //<THIS WILL BE THE FULLY QUALIFIED NAME OF THE DATA FROM THE CUSTOMER - WILL NEED TO RESIDE IN THE strInputJobFolder>;
               // string strSortedData = strPresortDataJobFolder + strJobNumber + " - Sorted.csv";
                string strOutputData = strPresortDataJobFolder + strJobNumber + " - Output.csv";
                string strMJBProcess = strTaskmasterJobsFolder + strProjectName + ".mjb";
                string strCommandLine = "-j \"" + Path.GetFileName(strMJBProcess) + "\" -u \"AUTO\" -w \"AUTO\" -r";
               // string strMailDatSettings = strMailDatSettingsFolder + strProjectName + ".mds";
                string strBCCListName = strJobNumber + " - " + strProjectName;
                string strBCCDatabase = strBCCJobFolder + strBCCListName + ".dbf";
                strInvalidsRemoved = strWorkingJobFolder + strJobNumber + " - Invalids Removed.xls";
                //strSortedData = strPresortDataJobFolder + strJobNumber + " - Sorted.csv";

                Process cmdProcess = new Process();
                ProcessStartInfo cmdProcessStartInfo = new ProcessStartInfo();

                StreamWriter streamMailDatSettings;
                StreamWriter streamMJBProcess;

                Console.WriteLine(string.Format("started created mds file."));


                //////////////////////////////////////////////////////////////////////////////////////////
                ////////////////////////////////////////////////////////// CREATING MAIL.DAT SETTINGS FILE
                //////////////////////////////////////////////////////////////////////////////////////////

                DataTable dtDistinct = new DataTable();
                dtDistinct.Columns.Add("OutputStream", typeof(string));
                DataRow dr = dtDistinct.NewRow();
                dr["OutputStream"] = "Label - Heeter";
                dtDistinct.Rows.Add(dr);

                DataRow dr1 = dtDistinct.NewRow();
                dr1["OutputStream"] = "Label - Highmark";
                dtDistinct.Rows.Add(dr1);

                List<string> lstSettingFiles = new List<string>();
                foreach (DataRow drDistinct in dtDistinct.Rows)
                {
                    string strSettingName = string.Format("Batch - {0}", drDistinct["OutputStream"].ToString());
                    string strMailDatSettings = Path.Combine(strMailDatSettingsFolder, string.Format("{0}.mds", strSettingName));
                    lstSettingFiles.Add(strMailDatSettings);

                    streamMailDatSettings = new StreamWriter(strMailDatSettings, false);

                    streamMailDatSettings.WriteLine("  FileSetName = '" + strJobNumber + "'");
                    streamMailDatSettings.WriteLine("  FileSource = 'Heeter Direct'");
                    streamMailDatSettings.WriteLine("  Location = '" + strDocsJobFolder + "'");
                    streamMailDatSettings.WriteLine("  JobNumber = '" + strJobNumber + "'");
                    streamMailDatSettings.WriteLine("  JobName = '" + strJobNumber + "'");
                    streamMailDatSettings.WriteLine("  ContactName = '" + strMailDatContactName + "'");
                    streamMailDatSettings.WriteLine("  ContactEmail = '" + strMailDatContactEmail + "'");
                    streamMailDatSettings.WriteLine("  ContactPhone = '" + strMailDatContactPhone + "'");
                    streamMailDatSettings.WriteLine("  SegmentingCriteria = '" + strJobNumber + "'");
                    streamMailDatSettings.WriteLine("  PostalOne = True");
                    streamMailDatSettings.WriteLine("  IMRTable = False");
                    streamMailDatSettings.WriteLine("  WSRTable = False");
                    streamMailDatSettings.WriteLine("  PDRTable = True");
                    streamMailDatSettings.WriteLine("  SFRTable = False");
                    streamMailDatSettings.WriteLine("  SNRTable = False");
                    streamMailDatSettings.WriteLine("  MSRTable = False");
                    streamMailDatSettings.WriteLine("  MIRTable = False");
                    streamMailDatSettings.WriteLine("  PBCTable = False");
                    streamMailDatSettings.WriteLine("  FixedBatch = False");
                    streamMailDatSettings.WriteLine("  BatchSize = 300");
                    streamMailDatSettings.WriteLine("  UseMailDatFolder = True");
                    streamMailDatSettings.WriteLine("  UseJobNumber = True");
                    streamMailDatSettings.WriteLine("  ZipDatabase = True");
                    streamMailDatSettings.WriteLine("  DeleteDatabaseAfterZipping = False");
                    streamMailDatSettings.WriteLine("  SegmentDescription = '" + strJobNumber + "'");
                    streamMailDatSettings.WriteLine("  MailingFacility = 'Pittsburgh PA'");
                    streamMailDatSettings.WriteLine("  MailingFacilityZIP4 = '152901001'");
                    streamMailDatSettings.WriteLine("  NumericDisplayContainerID = False");
                    streamMailDatSettings.WriteLine("  DetachedAddressLabels = False");
                    streamMailDatSettings.WriteLine("  UseConfirmBarcode = False");
                    streamMailDatSettings.WriteLine("  Services = False");
                    streamMailDatSettings.WriteLine("  ContainerInfo = True");
                    streamMailDatSettings.WriteLine("  EDocSenderCrid = '4057074'");
                    streamMailDatSettings.WriteLine("  BypassSeamlessAcceptance = False");
                    streamMailDatSettings.WriteLine("  USEIMpb = False");
                    streamMailDatSettings.WriteLine("  MailPieceName = 'FC'");
                    streamMailDatSettings.WriteLine("  Enclosure = False");
                    streamMailDatSettings.WriteLine("  RideAlong = False");
                    streamMailDatSettings.WriteLine("  RepositionableNote = False");
                    streamMailDatSettings.WriteLine("  ComponentDescription = '" + strJobNumber + "'");
                    streamMailDatSettings.WriteLine("  ContentOfMail = '  '");
                    streamMailDatSettings.WriteLine("  PostalPriceIncentiveType = '  '");
                    streamMailDatSettings.WriteLine("  EnclosureBulkInsurance = False");
                    streamMailDatSettings.WriteLine("  ActualEntryDiffers = False");
                    streamMailDatSettings.WriteLine("  InHomeRange = 0");
                    streamMailDatSettings.WriteLine("  EInduction = False");
                    streamMailDatSettings.WriteLine("  AcceptMisshipped = False");
                    streamMailDatSettings.WriteLine("  ContainerTags = True");
                    streamMailDatSettings.WriteLine("  MailerMailerLocation = 'Pittsburgh PA'");
                    streamMailDatSettings.WriteLine("  InformationLine = True");
                    streamMailDatSettings.WriteLine("  ResetUserInformationSackNumber = False");
                    streamMailDatSettings.WriteLine("  IMContainerTags = True");
                    streamMailDatSettings.WriteLine("  IMContainerTagsMailerID = '" + strHeeterMailerID + "'");
                    streamMailDatSettings.WriteLine("  Newspaper = False");
                    streamMailDatSettings.WriteLine("  PostagePaymentMethod = 'P'");
                    streamMailDatSettings.WriteLine("  ResetPackageID = False");
                    streamMailDatSettings.WriteLine("  SequentialPieceID = False");
                    streamMailDatSettings.WriteLine("  MailPieceStatus = False");
                    streamMailDatSettings.WriteLine("  AddACSKeylineCheckDigit = False");
                    streamMailDatSettings.WriteLine("  BulkInsurance = False");
                    streamMailDatSettings.WriteLine("  Confirm14DigitPlanetBarcode = False");
                    streamMailDatSettings.WriteLine("  IMBUseServicesExpression = False");
                    streamMailDatSettings.WriteLine("  IMBMailerIDExpression = '" + strHeeterMailerID + "'");
                    streamMailDatSettings.WriteLine("  IMBSerialNumberExpression = ''");
                    streamMailDatSettings.WriteLine("  IMBConfirmSampling = False");
                    streamMailDatSettings.WriteLine("  IMpbUseServicesExpression = False");
                    streamMailDatSettings.WriteLine("  IMpbZIPFormat = '9'");
                    streamMailDatSettings.WriteLine("  IMpbNonUSPSValid = False");
                    streamMailDatSettings.WriteLine("  InputMode = 'OVERWRITE'");
                    streamMailDatSettings.WriteLine("  FileSetStatus = 'O'");
                    streamMailDatSettings.WriteLine("  Version = '" + strMailDatVersion + "'");
                    streamMailDatSettings.WriteLine("  MoveUpdateDate = ''");
                    streamMailDatSettings.WriteLine("  WalkSequenceDate = ''");
                    streamMailDatSettings.WriteLine("  SubstitutedContainerPrep = ' '");
                    streamMailDatSettings.WriteLine("  MailPieceAdStatus = 'N'");
                    streamMailDatSettings.WriteLine("  ComponentRateType = 'R'");
                    streamMailDatSettings.WriteLine("  RideAlongProcessingCategory = 'LT'");
                    streamMailDatSettings.WriteLine("  RideAlongWeightStatus = 'F'");
                    streamMailDatSettings.WriteLine("  RideAlongWeightSource = 'C'");
                    streamMailDatSettings.WriteLine("  EnclosureType = 'N'");
                    streamMailDatSettings.WriteLine("  EnclosureRateType = 'X'");
                    streamMailDatSettings.WriteLine("  EnclosureProcessingCategory = 'LT'");
                    streamMailDatSettings.WriteLine("  EnclosureWeightStatus = 'F'");
                    streamMailDatSettings.WriteLine("  EnclosureWeightSource = 'C'");
                    streamMailDatSettings.WriteLine("  ActualEntryFacilityType = ' '");
                    streamMailDatSettings.WriteLine("  ContainerStatus = 'R'");
                    streamMailDatSettings.WriteLine("  ShipScheduledDateTime = ''");
                    streamMailDatSettings.WriteLine("  ShipDate = ''");
                    streamMailDatSettings.WriteLine("  InHomeDate = ''");
                    streamMailDatSettings.WriteLine("  StatementDateTime = ''");
                    streamMailDatSettings.WriteLine("  InductionDateTime = ''");
                    streamMailDatSettings.WriteLine("  InductionActualDateTime = ''");
                    streamMailDatSettings.WriteLine("  InternalDate = ''");
                    streamMailDatSettings.WriteLine("  PendingPeriodical = 'N'");
                    streamMailDatSettings.WriteLine("  ContainerChargeMethod = '2'");
                    streamMailDatSettings.WriteLine("  IssueDate = ''");
                    streamMailDatSettings.WriteLine("  ComponentAdStatus = 'N'");
                    streamMailDatSettings.WriteLine("  ServiceSet = ''");
                    streamMailDatSettings.WriteLine("  ConfirmService = '22'");
                    streamMailDatSettings.WriteLine("  PlanetBarcodeOption = 'S'");
                    streamMailDatSettings.WriteLine("  SeedType = 'R'");
                    streamMailDatSettings.WriteLine("  ZIP4EncodingDate = ''");
                    streamMailDatSettings.WriteLine("  PickupScheduledDateTime = ''");
                    streamMailDatSettings.WriteLine("  PickupDateTime = ''");
                    streamMailDatSettings.WriteLine("  MoveUpdateMethod = '0'");
                    streamMailDatSettings.WriteLine("  USPSPickupMailing = 'N'");
                    streamMailDatSettings.WriteLine("  USPSPickup = 'N'");
                    streamMailDatSettings.WriteLine("  SeamlessAcceptanceIndicator = ' '");
                    streamMailDatSettings.WriteLine("  FullServiceParticipation = 'F'");
                    streamMailDatSettings.WriteLine("  SASPPreparationOption = ' '");
                    streamMailDatSettings.WriteLine("  ACSKeyline = 'N'");
                    streamMailDatSettings.WriteLine("  DetachedMailingLabel = ' '");
                    streamMailDatSettings.WriteLine("  CharacteristicFee = '  '");
                    streamMailDatSettings.WriteLine("  PostagePaymentOption = 'D'");

                    streamMailDatSettings.Close();
                    streamMailDatSettings.Dispose();


                }

                //////////////////////////////////////////////////////////////////////////////////////////
                //////////////////////////////////////////////////////////////////////// CREATING MJB FILE
                //////////////////////////////////////////////////////////////////////////////////////////
                streamMJBProcess = new StreamWriter(strMJBProcess, false);

                streamMJBProcess.WriteLine("[NEWLISTTEMPLATE-1]");
                streamMJBProcess.WriteLine("DESCRIPTION=\"" + strBCCListName + "\"");
                streamMJBProcess.WriteLine("FILENAME=\"" + strBCCDatabase + "\"");
                streamMJBProcess.WriteLine("OVERWRITE=Y");
                streamMJBProcess.WriteLine("GROUP=\"AUTO\"");
                streamMJBProcess.WriteLine("SETTINGS=\"Highmark Fulfillment Template\"");
                streamMJBProcess.WriteLine("USEINDEXES=Y");
                streamMJBProcess.WriteLine("");

                streamMJBProcess.WriteLine("[IMPORT-2]");
                streamMJBProcess.WriteLine("LIST=\"" + strBCCDatabase + "\"");
                streamMJBProcess.WriteLine("SETTINGS=\"" + strImportName + "\"");
                streamMJBProcess.WriteLine("FILENAME=\"" + strInputDataName + "\"");
                streamMJBProcess.WriteLine("");

                streamMJBProcess.WriteLine("[MODIFY-3]");
                streamMJBProcess.WriteLine("LIST=\"" + strBCCDatabase + "\"");
                streamMJBProcess.WriteLine("SETTINGS=\"Standard Address Block Filter\"");
                streamMJBProcess.WriteLine("SELECTIVITY=NONE");
                streamMJBProcess.WriteLine("");

                streamMJBProcess.WriteLine("[ENCODE-4]");
                streamMJBProcess.WriteLine("LIST=\"" + strBCCDatabase + "\"");
                streamMJBProcess.WriteLine("SELECTIVITY=NONE");
                streamMJBProcess.WriteLine("ADDRESSGROUPS=\"MAIN\"");
                streamMJBProcess.WriteLine("SWAP=Y");
                streamMJBProcess.WriteLine("STANDARDIZEADDRESS=Y");
                streamMJBProcess.WriteLine("STANDARDIZECITY=Y");
                streamMJBProcess.WriteLine("ABBREVIATECITY=N");
                streamMJBProcess.WriteLine("IGNORENONUSPS=Y");
                streamMJBProcess.WriteLine("EXTENDEDMATCHING=N");
                streamMJBProcess.WriteLine("CASE=\"ASIS\"");
                streamMJBProcess.WriteLine("FIRMASIS=Y");
                streamMJBProcess.WriteLine("ZIP5CHECKDIGIT=N");
                streamMJBProcess.WriteLine("SUMMARYPAGE=N");
                streamMJBProcess.WriteLine("NDIREPORT=N");
                streamMJBProcess.WriteLine("COPIES=0");
                streamMJBProcess.WriteLine("PREPAREDFOR=NONE");
                streamMJBProcess.WriteLine("");

                //if (iInputRecords > 99)
                //{
                //    streamMJBProcess.WriteLine("[DATASERVICES-5]");
                //    streamMJBProcess.WriteLine("LIST=\"" + strBCCDatabase + "\"");
                //    streamMJBProcess.WriteLine("PROCESSES=\"FSP\"");
                //    streamMJBProcess.WriteLine("CLASSOFMAIL=\"F\"");
                //    streamMJBProcess.WriteLine("MAILINGZIPCODE=\"15290\"");
                //    streamMJBProcess.WriteLine("LISTOWNER=\"" + strNCOACompany + "\"");
                //    streamMJBProcess.WriteLine("PREPAID=Y");
                //    streamMJBProcess.WriteLine("EXTENDEDMATCHING=Y");
                //    streamMJBProcess.WriteLine("PAFELECTRONIC=N");
                //    streamMJBProcess.WriteLine("JOBPASSWORD=120000001EE18648DBD198CC853206D2E93C8604120DFBA79417C9E2A68A8C523881F5D30316FA979FC95A5484F12977A9569A570F338049AB987585C4609D62677D8B901FC0672C8E037EC8212AEEC9D3BB8E0E");
                //    streamMJBProcess.WriteLine("ADDRESSGROUPS=\"MAIN\"");
                //    streamMJBProcess.WriteLine("ORDERTERMSACCEPTED=Y");
                //    streamMJBProcess.WriteLine("CASE=\"AUTO\"");
                //    streamMJBProcess.WriteLine("COAAUDITEXPORT=Y");
                //    streamMJBProcess.WriteLine("STANDARDIZECITY=N");
                //    streamMJBProcess.WriteLine("HIDERETURNCODES=\"10 11 12 13 17 21 26 27 28 33 98\"");
                //    streamMJBProcess.WriteLine("");
                //}
                //else
                //{
                    streamMJBProcess.WriteLine("[HIDE-5]");
                    streamMJBProcess.WriteLine("LIST=\"" + strBCCDatabase + "\"");
                    streamMJBProcess.WriteLine("SELECTIVITYEXPRESSION=\"([RC] <> '22') AND ([RC] <> '31') AND ([RC] <> '32')\"");
                    streamMJBProcess.WriteLine("");
                //}

                streamMJBProcess.WriteLine("[EXPORT-6]");
                streamMJBProcess.WriteLine("LIST=\"" + strBCCDatabase + "\"");
                streamMJBProcess.WriteLine("FILENAME=\"" + strInvalidsRemoved + "\"");
                streamMJBProcess.WriteLine("SETTINGS=\"Highmark Fulfillment Export\"");
                streamMJBProcess.WriteLine("SELECTIVITY=\"Hidden Record\"");
                streamMJBProcess.WriteLine("INDEX=NONE");
                streamMJBProcess.WriteLine("OVERWRITE=Y");
                streamMJBProcess.WriteLine("");

                streamMJBProcess.WriteLine("[DELETEHIDDENRECORDS-7]");
                streamMJBProcess.WriteLine("LIST=\"" + strBCCDatabase + "\"");
                streamMJBProcess.WriteLine("");

                streamMJBProcess.WriteLine("[MODIFY-8]");
                streamMJBProcess.WriteLine("LIST=\"" + strBCCDatabase + "\"");
                streamMJBProcess.WriteLine("SETTINGS=\"CASING: Mixed Case Address Block\"");
                streamMJBProcess.WriteLine("SELECTIVITY=NONE");
                streamMJBProcess.WriteLine("");

                streamMJBProcess.WriteLine("[PRESORT-9]");
                streamMJBProcess.WriteLine("LIST=\"" + strBCCDatabase + "\"");
                streamMJBProcess.WriteLine("ADDRESSGROUP=\"MAIN\"");
                streamMJBProcess.WriteLine("SELECTIVITY=NONE");
                streamMJBProcess.WriteLine("SETTINGS=\"" + strPresortSettings + "\"");
                streamMJBProcess.WriteLine("FULLSERVICEPARTICIPATION=\"FULLSERVICE\"");
                streamMJBProcess.WriteLine(" ");

                int itaskNbr = 10;
                for (int index = 0; index < dtDistinct.Rows.Count; index++)
                {
                    string strSortedData = Path.Combine(strPresortDataJobFolder, string.Format("Batch - {0}_OUTPUT_DATA.csv", dtDistinct.Rows[index]["OutputStream"].ToString()));

                    streamMJBProcess.WriteLine(string.Format("[PRESORTEDLABELS-{0}]", itaskNbr));
                    streamMJBProcess.WriteLine("LIST=\"" + strBCCDatabase + "\"");
                    streamMJBProcess.WriteLine("SETTINGS=\"" + strOutputLabel + "\"");
                    streamMJBProcess.WriteLine("PRESORTNAME=\"" + string.Format("Batch - {0}", dtDistinct.Rows[index]["OutputStream"].ToString()) + "\"");
                    streamMJBProcess.WriteLine("STREAMLIST=\"MERGED;AUTO/NONAUTO;AUTO;MACH;SINGLE PC\"");
                    streamMJBProcess.WriteLine("ABSOLUTECONTAINERNUMBERS=Y");
                    streamMJBProcess.WriteLine("FILENAME=\"" + strSortedData + "\"");
                    streamMJBProcess.WriteLine("OVERWRITE=Y");
                    streamMJBProcess.WriteLine(" ");

                    itaskNbr++;

                    streamMJBProcess.WriteLine(string.Format("[MAILDAT-{0}]", itaskNbr));
                    streamMJBProcess.WriteLine("LIST=\"" + strBCCDatabase + "\"");
                    streamMJBProcess.WriteLine("SETTINGS=\"" + Path.GetFileNameWithoutExtension(lstSettingFiles[index]) + "\"");
                    streamMJBProcess.WriteLine("PRESORTNAME=\"" + string.Format("Batch - {0}", dtDistinct.Rows[index]["OutputStream"].ToString()) + "\"");
                    streamMJBProcess.WriteLine("STREAMLIST=\"MERGED;AUTO/NONAUTO;AUTO;MACH;SINGLE PC\"");
                    streamMJBProcess.WriteLine(" ");

                    itaskNbr++;
                }
                //streamMJBProcess.WriteLine("[EXPORT-12]");
                //streamMJBProcess.WriteLine("LIST=\"" + strBCCDatabase + "\"");
                //streamMJBProcess.WriteLine("FILENAME=\"" + strOutputData + "\"");
                //streamMJBProcess.WriteLine("SETTINGS=\"Mylan EPCC Export\"");
                //streamMJBProcess.WriteLine("SELECTIVITY=NONE");
                //streamMJBProcess.WriteLine("INDEX=NONE");
                //streamMJBProcess.WriteLine("OVERWRITE=Y");
                //streamMJBProcess.WriteLine("");

                streamMJBProcess.WriteLine(string.Format("[TERMINATE-{0}]", itaskNbr));
                streamMJBProcess.WriteLine("");

                streamMJBProcess.Close();
                streamMJBProcess.Dispose();



                //////////////////////////////////////////////////////////////////////////////////////////
                ////////////////////////////////////////////////////////////// RUNNING FIRST PASS MJB FILE
                //////////////////////////////////////////////////////////////////////////////////////////
                cmdProcessStartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                cmdProcessStartInfo.RedirectStandardError = false;
                cmdProcessStartInfo.RedirectStandardOutput = false;
                cmdProcessStartInfo.UseShellExecute = true;
                cmdProcessStartInfo.CreateNoWindow = true;
                cmdProcessStartInfo.FileName = strBCCMailManEXE;
                cmdProcessStartInfo.Arguments = strCommandLine;



                cmdProcess = Process.Start(cmdProcessStartInfo);
                cmdProcess.WaitForExit();
                cmdProcess.Dispose();

                Console.WriteLine(string.Format("BCC Process executed successfully."));


                //File.Delete(strMJBProcess);
                //File.Delete(strMailDatSettings);
            }
            catch (Exception exception)
            {
                //LogFile(exception.ToString(), true);

                Console.WriteLine(string.Format("Exception Occured: {0}.", exception.Message));

                return false;
            }

            return true;
        }

        #endregion
    }
}

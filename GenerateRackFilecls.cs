using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Common;
using DAL;
using LSEXT;
using LSSERVICEPROVIDERLib;
using One1.Controls;

namespace GenerateRackFile
{


    [ComVisible(true)]
    [ProgId("GenerateRackFile.GenerateRackFilecls")]
    public class GenerateRackFile : IEntityExtension
    {
        private INautilusServiceProvider sp;

        public ExecuteExtension CanExecute(ref IExtensionParameters Parameters)
        {
            return ExecuteExtension.exEnabled;
        }

        public void Execute(ref LSExtensionParameters Parameters)
        {
            try
            {
                //Generate Rack File For BAX

                sp = Parameters["SERVICE_PROVIDER"];
                var ntlsCon = Utils.GetNtlsCon(sp);
                Utils.CreateConstring(ntlsCon);
                var records = Parameters["RECORDS"];

                var rackId = records.Fields["RACK_ID"].Value;

                long id = (long.Parse(rackId.ToString()));

                var dal = new DataLayer();

                dal.Connect();


                var conversionPhrase = dal.GetPhraseByName("short name to BAX conversion");
                if (conversionPhrase == null) throw new ArgumentNullException("Phrase doesn't exist");
                var conversionTable = conversionPhrase.PhraseEntries.ToList();
                //Get rack by id
                var currentRack = dal.GetRackById(id);
                var rackUsage = currentRack.RACK_USAGE.FirstOrDefault(x => x.IS_CURRENT == "T");
                if (rackUsage == null) throw new ArgumentNullException("rackUsage is empty");
                var ra = rackUsage.RACK_ALIQUOT.OrderBy(x => x.ABSOLUTE_POSITION);


                List<FileDetails> list = new List<FileDetails>();
                foreach (RACK_ALIQUOT rackAliquot in ra)
                {


                    var currentAliquot = rackAliquot.ALIQUOT;

                    var filedetails = new FileDetails();
                    filedetails.AutoNumber = rackAliquot.ABSOLUTE_POSITION;
                    filedetails.AliquotId = currentAliquot.AliquotId;
                    filedetails.AliquotName = currentAliquot.Name;
                    if (string.IsNullOrEmpty(currentAliquot.ShortName))
                    {
                        throw new ArgumentNullException("ShortName  doesn't exist for" + currentAliquot.Name);

                    }
                    var c =
                        conversionTable.FirstOrDefault(x => x.PhraseName == currentAliquot.ShortName);
                    if (c == null || string.IsNullOrEmpty(c.PhraseDescription))
                    {
                        throw new ArgumentNullException("Not conversion for " + currentAliquot.ShortName);
                    }
                    filedetails.BaxName = c.PhraseDescription;
                    list.Add(filedetails);
                }

                //Get location to save files
                var p = dal.GetPhraseByName("Location folders");
                var path = p.PhraseEntries.FirstOrDefault(x => x.PhraseDescription == "Rack files to BAX");


                //Write to file
                string fullPath = path.PhraseName + "BAX-RACK-" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".bax";



                using (FileStream file = new FileStream(fullPath, FileMode.Append, FileAccess.Write))
                {
                    var streamWriter = new StreamWriter(file);
                    streamWriter.WriteLine("Well Info	Sample ID	Description	Sample Type");
                    foreach (var fd in list)
                    {
                        streamWriter.WriteLine("{0}\t{1}\t{2}\t{3}\t", fd.AutoNumber, fd.AliquotId, fd.AliquotName, fd.BaxName);
                    }
                    streamWriter.WriteLine("{0}\t", "96");

                    streamWriter.Close();

                }

                CustomMessageBox.Show("הקובץ הופק בהצלחה.", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            catch (Exception e)
            {
                Logger.WriteLogFile(e);
                CustomMessageBox.Show("Error" + e.Message, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }


        }
    }

    class FileDetails
    {
        public int AutoNumber { get; set; }
        public long? AliquotId { get; set; }
        public string AliquotName { get; set; }
        public string BaxName { get; set; }
    }
}

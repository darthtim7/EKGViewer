using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace EKGViewer.File
{
    public class WFDBData
    {

        private string Filename;
        private WFDBHeader? Header;
        private List<WFDBSignal> Signals;

        public WFDBData(string filename)
        {
            Filename = filename;
            Header = null;
            Signals = new List<WFDBSignal>();
        }

        public async Task Load()
        {
            try
            {
                var reader = new WFDBReader("Ekg");

                Header = await reader.ReadHeaderAsync(Filename);

                for (var i = 0; i < Header.Signals.Count; i++)
                {
                    var sigRead = await reader.LoadSignalAsync(Filename, i);
                    Signals.Add(sigRead);
                }
            }
            catch (Exception)
            {
                // TODO: handle exceptions here
                throw;
            }
        }
    }
}
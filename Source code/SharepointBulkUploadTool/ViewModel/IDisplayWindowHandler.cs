using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharepointBulkUploadTool.ViewModel
{
    public interface IDisplayWindowHandler
    {
        void UpdateGridColumns(object dataContext);

        void ShowMessage(string message, string header = "");

        void ShowErrorMessage(string message, string header = "");
    }
}

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using VSTOAddonForPowerPoint.Properties;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace VSTOAddonForPowerPoint
{
    [ComVisible(true)]
    public class RibbonController : Microsoft.Office.Core.IRibbonExtensibility
    {
        private Microsoft.Office.Core.IRibbonUI _ribbonUi;

        public string GetCustomUI(string ribbonID) =>
            @"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
                        <ribbon>
                           <tabs>
                                <tab id='sample_tab' label='Custom button'>
                                    <group id='sample_group' label='Operations'>
                                        <button id='do_1' label='Do diagram' size='large' onAction='OnClick'/>
                                    </group>
                                </tab>
                            </tabs>
                        </ribbon>
                    </customUI>";

        public void OnLoad(Microsoft.Office.Core.IRibbonUI ribbonUI)
        {
            _ribbonUi = ribbonUI;
        }

        public async void OnClick(Microsoft.Office.Core.IRibbonControl control)
        {
            PowerPoint.Slide currentSlide = Globals.ThisAddIn.Application.ActivePresentation.Slides[1];
            var chart = currentSlide.Shapes.AddChart(Office.XlChartType.xlColumnClustered, 0, 0, 500, 500).Chart;           
            chart.ChartData.Activate();

            if (chart.HasLegend)  chart.Legend.Font.Bold= true;

            chart.ApplyDataLabels(XlDataLabelsType.xlDataLabelsShowValue);
            
        }
        
    }
}

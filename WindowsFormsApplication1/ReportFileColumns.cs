using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApplication1
{
   public  class ReportFileColumns
    {
        ///<summary>
        ///*****办事处**********
        /// </summary>
        public string office = "南宁办事处";
        ///<summary>
        ///*****产品型号*********
        /// </summary>
        public string Model="";
        ///<summary>
        /// ****整机编号*********
        /// </summary>
        public string ProductID="";
        ///<summary>
        /// ****部件编号********（可以为空）
        /// </summary>
        public string ComponentID="";
        ///<summary>
        /// ***出厂日期********
        /// </summary>
        public string ProductionDate="";
        ///<summary>
        /// ****故障描述*******
        /// </summary>
        public string ErrorSrc="";
        ///<summary>
        /// ****处理方法*******
        /// </summary>
        public string DealSrc="";
        ///<summary>
        ///*****更换部件******* 
        /// </summary>
        public string ReplacedComponent="无";
        ///<summary>
        /// *****用户名称******
        /// </summary>
        public string ClientName="";
        ///<summary>
        /// *****维修人********
        /// </summary>
        public string Engineer="";
        ///<summary>
        /// *****维修日期******
        /// </summary>
        public string MaintenanceDate="";
        ///<summary>
        /// *****派工单号******
        /// </summary>
        public string PaiID="";
        ///<summary>
        /// *****维修单编号****
        /// </summary>
        public string ServiceID="";

        
        
    }
}

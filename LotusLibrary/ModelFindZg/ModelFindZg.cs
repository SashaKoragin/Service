using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LotusLibrary.ModelFindZg
{
    public class ModelFindZg
    {
        private string fioFindMemoField;

        private string fioFindLotusField;

        private string innField;

        private string zgNumberField;

        private DateTime inCard_RgDateField;

        private string ex_ExecDirectField;

        private string deptField;

        private string outNumberField;

        private Nullable<DateTime>checkUpDateField;

        private string numberField;

        public Nullable<DateTime> CheckUpDate
        {
            get
            {
                return this.checkUpDateField;
            }
            set
            {
                this.checkUpDateField = value;
            }
        }
        
        public string Number
        {
            get
            {
                return this.numberField;
            }
            set
            {
                this.numberField = value;
            }
        }

        public string OutNumber
        {
            get
            {
                return this.outNumberField;
            }
            set
            {
                this.outNumberField = value;
            }
        }

        public string Dept
        {
            get
            {
                return this.deptField;
            }
            set
            {
                this.deptField = value;
            }
        }


        public string FioFindMemo
        {
            get
            {
                return this.fioFindMemoField;
            }
            set
            {
                this.fioFindMemoField = value;
            }
        }
        public string FioFindLotus
        {
            get
            {
                return this.fioFindLotusField;
            }
            set
            {
                this.fioFindLotusField = value;
            }
        }


        public string Inn
        {
            get
            {
                return this.innField;
            }
            set
            {
                this.innField = value;
            }
        }

        public string ZgNumber
        {
            get
            {
                return this.zgNumberField;
            }
            set
            {
                this.zgNumberField = value;
            }
        }

        public DateTime InCard_RgDate
        {
            get
            {
                return this.inCard_RgDateField;
            }
            set
            {
                this.inCard_RgDateField = value;
            }
        }

        public string Ex_ExecDirect
        {
            get
            {
                return this.ex_ExecDirectField;
            }
            set
            {
                this.ex_ExecDirectField = value;
            }
        }
    }
}

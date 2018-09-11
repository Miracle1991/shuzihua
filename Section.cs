using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace shuzihua
{
    class Section
    {
        private string _contour_min;
        private string _contour_max;
        private string _cv_cont_min;
        private string _cv_cont_max;
        private string _cc_cont_min;
        private string _cc_cont_max;
        private string _le_contr_min;
        private string _le_contr_max;
        private string _te_contr_min;
        private string _te_contr_max;
        private string _stack_x;
        private string _stack_y;
        private string _twist_ang;
        private string _te_redius;
        private string _le_redius;
        private string _le_aq;
        private string _chord_wid;
        private string _b;
        private string _b1;
        private string _extreme;
        private string _le_1_5;
        private string _te_1_5;
        private string _qiexian;


        public string STACK_X
        {
            get
            {
                return _stack_x;
            }
            set
            {
                _stack_x = value;
            }
        }

        public string STACK_Y
        {
            get
            {
                return _stack_y;
            }
            set
            {
                _stack_y = value;
            }
        }

        public string TWIST_ANG
        {
            get
            {
                return _twist_ang;
            }
            set
            {
                _twist_ang = value;
            }
        }

        public string CC_CONT_MIN
        {
            get
            {
                return _cc_cont_min;
            }
            set
            {
                _cc_cont_min = value;
            }
        }

        public string CC_CONT_MAX
        {
            get
            {
                return _cc_cont_max;
            }
            set
            {
                _cc_cont_max = value;
            }
        }

        public string CV_CONT_MAX
        {
            get
            {
                return _cv_cont_max;
            }
            set
            {
                _cv_cont_max = value;
            }
        }

        public string CV_CONT_MIN
        {
            get
            {
                return _cv_cont_min;
            }
            set
            {
                _cv_cont_min = value;
            }
        }

        public string LE_CONTR_MAX
        {
            get
            {
                return _le_contr_max;
            }
            set
            {
                _le_contr_max = value;
            }
        }

        public string LE_CONTR_MIN
        {
            get
            {
                return _le_contr_min;
            }
            set
            {
                _le_contr_min = value;
            }
        }

        public string TE_CONTR_MAX
        {
            get
            {
                return _te_contr_max;
            }
            set
            {
                _te_contr_max = value;
            }
        }

        public string TE_CONTR_MIN
        {
            get
            {
                return _te_contr_min;
            }
            set
            {
                _te_contr_min = value;
            }
        }
        public string CONTOUR_MIN
        {
            get
            {
                return _contour_min;
            }
            set
            {
                _contour_min = value;
            }
        }

        public string CONTOUR_MAX
        {
            get
            {
                return _contour_max;
            }
            set
            {
                _contour_max = value;
            }
        }

        public string Le_aq
        {
            get
            {
                return _le_aq;
            }
            set
            {
                _le_aq = value;
            }
        }


        public string Te_redius
        {
            get
            {
                return _te_redius;
            }
            set
            {
                _te_redius = value;
            }
        }

        public string Le_redius
        {
            get
            {
                return _le_redius;
            }
            set
            {
                _le_redius = value;
            }
        }

        public string Chord_wid
        {
            get
            {
                return _chord_wid;
            }
            set
            {
                _chord_wid = value;
            }
        }

        public string B
        {
            get
            {
                return _b;
            }
            set
            {
                _b = value;
            }
        }

        public string B1
        {
            get
            {
                return _b1;
            }
            set
            {
                _b1 = value;
            }
        }

        public string Extreme
        {
            get
            {
                return _extreme;
            }
            set
            {
                _extreme = value;
            }
        }

        public string Le_1_5
        {
            get
            {
                return _le_1_5;
            }
            set
            {
                _le_1_5 = value;
            }
        }

        public string Te_1_5
        {
            get
            {
                return _te_1_5;
            }
            set
            {
                _te_1_5 = value;
            }
        }

        public string Qiexian
        {
            get
            {
                return _qiexian;
            }
            set
            {
                _qiexian = value;
            }
        }


        public Section()
        {
            this._contour_min = "";
            this._contour_max = "";
            this._cv_cont_min = "";
            this._cv_cont_max = "";
            this._cc_cont_min = "";
            this._cc_cont_max = "";
            this._le_contr_min = "";
            this._le_contr_max = "";
            this._te_contr_min = "";
            this._te_contr_max = "";
            this._stack_x = "";
            this._stack_y = "";
            this._twist_ang = "";
            this._le_redius = "";
            this._te_redius = "";
            this._le_aq = "";
            this._chord_wid = "";
            this._b1 = "";
            this._b = "";
            this._extreme = "";
            this._le_1_5 = "";
            this._te_1_5 = "";
            this._qiexian = "";


        }
    }
}




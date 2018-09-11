using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace shuzihua
{
    class myPoint
    {
        private List<string> _x;
        private List<string> _y;
        private List<string> _z;
        private List<string> _t;
        private List<string> _id;
        private List<int> _index;

        private bool x_valid;
        private bool y_valid;
        private bool z_valid;
        private bool t_valid;

        private List<bool> is_ok;
        private string _name;
        private int _col;
        private int cnt;

        private bool neg;

        public int Cnt
        {
            get
            {
                return cnt;
            }
            set
            {
                cnt = value;
            }
        }

        public List<bool> Is_ok
        {
            get
            {
                return is_ok;
            }
            set
            {
                is_ok = value;
            }
        }

        public List<string> X
        {
            get
            {
                return _x;
            }
            set
            {
                _x = value;
            }
        }

        public List<string> Y
        {
            get
            {
                return _y;
            }
            set
            {
                _y = value;
            }
        }

        public List<string> Z
        {
            get
            {
                return _z;
            }
            set
            {
                _z = value;
            }
        }

        public List<string> T
        {
            get
            {
                return _t;
            }
            set
            {
                _t = value;
            }
        }

        public List<string> ID
        {
            get
            {
                return _id;
            }
            set
            {
                _id = value;
            }
        }

        public List<int> Index
        {
            get
            {
                return _index;
            }
            set
            {
                _index = value;
            }
        }

        public string Name
        {
            get
            {
                return _name;
            }
            set
            {
                _name = value;
            }
        }

        public int Col
        {
            get
            {
                return _col;
            }
            set
            {
                _col = value;
            }
        }
        public bool Z_valid
        {
            get
            {
                return z_valid;
            }
            set
            {
                z_valid = value;
            }
        }
        public bool X_valid
        {
            get
            {
                return x_valid;
            }
            set
            {
                x_valid = value;
            }
        }
        public bool Y_valid
        {
            get
            {
                return y_valid;
            }
            set
            {
                y_valid = value;
            }
        }

        public bool T_valid
        {
            get
            {
                return t_valid;
            }
            set
            {
                t_valid = value;
            }
        }

        public bool Neg
        {
            get
            {
                return this.neg;
            }
            set
            {
                this.neg = value;
            }
        }

        public myPoint()
        {
            this._x = new List<string>();
            this._y = new List<string>();
            this._z = new List<string>();
            this._t = new List<string>();
            this._id = new List<string>();
            this._index = new List<int>();
            this.is_ok = new List<bool>();

            this.x_valid = false;
            this.y_valid = false;
            this.z_valid = false;
            this.t_valid = false;
            this.cnt = 0;
            
            this._name = "";
            this._col = 0;
            this.neg = false;
        }
    }
}

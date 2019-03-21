using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CommonLibrary.Exceptions
{
    public class NoSuchControlTypeFound:Exception
    {
        public NoSuchControlTypeFound(string msg) : base(msg)
        {

        }
    }
    public class NoSuchModuleFound:Exception
    {
        public NoSuchModuleFound(string msg) : base(msg)
        {

        }
    }
    public class NoSuchOperationFound : Exception
    {
        public NoSuchOperationFound(string msg) : base(msg)
        {

        }
    }
    public class NoSuchWindowTypeFound : Exception
    {
        public NoSuchWindowTypeFound(string msg) : base(msg)
        {

        }
    }
}

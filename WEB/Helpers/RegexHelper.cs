using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Helpers
{
    public class RegexHelper
    {
        public static string Alpha { get { return @"^[a-zA-Z]+$"; } }
        public static string AlphaNumeric { get { return @"^[a-zA-Z0-9]+$"; } }
        public static string Color { get { return @"^#?([a-fA-F0-9]{6}|[a-fA-F0-9]{3})$"; } }
        public static string Domain { get { return @"^([a-zA-Z0-9]([a-zA-Z0-9\-]{0,61}[a-zA-Z0-9])?\.)+[a-zA-Z]{2,8}$"; } }
        public static string Email { get { return @"^[a-zA-Z0-9.!#$%&'*+\/=?^_`{|}~-]+@[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?(?:\.[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?)*$"; } }
        public static string Integer { get { return @"^[-+]?\d+$"; } }
        public static string Number { get { return @"^[-+]?\d*(?:[\.\,]\d+)?$"; } }
    }
}

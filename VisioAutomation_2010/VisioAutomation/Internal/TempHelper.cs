﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace VisioAutomation.Internal
{
    internal class TempHelper
    {
        public static void _enforce_valid_result_type(System.Type result_type)
        {
            if (!_is_valid_result_type(result_type))
            {
                string msg = string.Format("Unsupported Result Type: {0}", result_type.Name);
                throw new Exceptions.InternalAssertionException(msg);
            }
        }

        public static bool _is_valid_result_type(System.Type result_type)
        {
            return (result_type == typeof(int)
                    || result_type == typeof(double)
                    || result_type == typeof(string));
        }

    }
}

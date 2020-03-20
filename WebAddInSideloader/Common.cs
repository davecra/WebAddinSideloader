using System;

namespace WebAddInSideloader
{
    public static class Common
    {
        /// <summary>
        /// Global error handler
        /// </summary>
        /// <param name="PobjEx"></param>
        /// <returns></returns>
        public static void HandleException(this Exception PobjEx, bool PbolNotify = true, string PstrMessage = "")
        {
            if(PbolNotify)
            {
                Console.WriteLine(PobjEx.Message + " " + PstrMessage);
            }
            else
            {
                // we do nothing 
            }
        }

        /// <summary>
        /// Passes the exception up. Use this with a throw statement
        /// </summary>
        /// <param name="PobjEx"></param>
        /// <param name="PstrMessage"></param>
        /// <returns></returns>
        public static Exception PassException(this Exception PobjEx, string PstrMessage)
        {
            return new Exception(PstrMessage, PobjEx);
        }
    }
}

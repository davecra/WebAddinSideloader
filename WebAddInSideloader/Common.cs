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
    }
}

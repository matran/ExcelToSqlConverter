using System.Data.SqlClient;
using System.Reflection;

namespace excel_file_loader_fm
{
    static class Class1
    {
        static FieldInfo rowsCopiedField = null;

        /// <summary>
        /// Gets the rows copied from the specified SqlBulkCopy object
        /// </summary>
        /// <param name="bulkCopy">The bulk copy.</param>
        /// <returns></returns>
        public static int GetRowsCopied(SqlBulkCopy bulkCopy)
        {
            if (rowsCopiedField == null)
            {
                rowsCopiedField = typeof(SqlBulkCopy).GetField("_rowsCopied", BindingFlags.NonPublic | BindingFlags.GetField | BindingFlags.Instance);
            }

            return (int)rowsCopiedField.GetValue(bulkCopy);
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.Access;
using Microsoft.Office.Interop.Access.Dao;

namespace Com.Ericmas001.Office.Access
{
    public static class AccessApplicationExtensions
    {
        public static IEnumerable<Dictionary<string, dynamic>> Select(this Application app, string selectQuery)
        {
            List<Dictionary<string, dynamic>> results = new List<Dictionary<string, dynamic>>();

            var rs = app.CurrentDb().OpenRecordset(selectQuery);
            while (!rs.EOF)
            {
                results.Add(rs.Fields.OfType<Field>().ToDictionary(field => field.Name, field => field.Value is DBNull ? null : field.Value));
                rs.MoveNext();
            }

            return results;
        }
        public static Dictionary<string, dynamic> SelectOne(this Application app, string selectQuery)
        {
            return app.Select(selectQuery).FirstOrDefault();
        }
        public static bool Update(this Application app, string updateQuery)
        {
            try
            {
                app.CurrentDb().Execute(updateQuery);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
    }
}

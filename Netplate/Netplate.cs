using System.Collections.Generic;

using Dataset = System.Collections.Generic.IDictionary<string, object>;

namespace Netplate
{
    public static class Netplate
    {
        /// <summary>
        /// Creates a single Mail Merge based on the template and data.
        /// </summary>
        /// <param name="template">The template used to insert the data into.</param>
        /// <param name="data">The data to insert into the template. Key for the Merge Field, Value for the content. Calls .ToString() on the value.</param>
        /// <returns>The a completed Mail Merge.</returns>
        public static string MailMerge(string template, Dataset data)
        {
            string result = template;

            ReplaceIfMergeFields(ref result, data);
            ReplaceBasicMergeFields(ref result, data);

            return result;
        }

        /// <summary>
        /// Creates multiple Mail Merges based on the template and data.
        /// </summary>
        /// <param name="template">The template used to insert the data into.</param>
        /// <param name="data">The data to insert into the template. Key for the Merge Field, Value for the content. Calls .ToString() on the value.</param>
        /// <returns>The completed Mail Merge in the same order as the data.</returns>
        public static IList<string> Create(string template, IList<Dataset> data)
        {
            List<string> results = new List<string>(data.Count);

            foreach(Dataset item in data)
            {
                results.Add(MailMerge(template, item));
            }

            return results;
        }

        private static void ReplaceBasicMergeFields(ref string text, Dataset data)
        {

        }

        private static void ReplaceIfMergeFields(ref string text, Dataset data)
        {

        }

        private static void DisplayLeftConditionalContent(ref string text, int start, bool isDisplayed)
        {

        }
    }
}

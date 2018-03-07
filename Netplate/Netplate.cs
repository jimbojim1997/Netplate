using System.Collections.Generic;

using Dataset = System.Collections.Generic.IDictionary<string, object>;

namespace Netplate
{
    public static class MailMerge
    {
        private const char MERGE_FIELD_MARKER_LEFT = '[';
        private const char MERGE_FIELD_MARKER_RIGHT = ']';
        private const char MERGE_FIELD_MARKER_CONDITIONAL = '~';
        private const string MERGE_FIELD_MARKER_IF = "if";
        private const string MERGE_FIELD_IF_OPERATOR_EQUAL = "==";
        private const string MERGE_FIELD_IF_OPERATOR_NOT_EQUAL = "!=";
        private const string MERGE_FIELD_IF_OPERATOR_GREATER_THAN = ">";
        private const string MERGE_FIELD_IF_OPERATOR_LESS_THAN = "<";
        private const string MERGE_FIELD_IF_OPERATOR_GREATER_OR_EQUAL = ">=";
        private const string MERGE_FIELD_IF_OPERATOR_LESS_OR_EQUAL = "<=";

        /// <summary>
        /// Creates a single Mail Merge based on the template and data.
        /// </summary>
        /// <param name="template">The template used to insert the data into.</param>
        /// <param name="data">The data to insert into the template. Key for the Merge Field, Value for the content. Calls .ToString() on the value.</param>
        /// <returns>The a completed Mail Merge.</returns>
        public static string Merge(string template, Dataset data)
        {
            string result = template;

            MergeIfMergeFields(ref result, data);
            MergeBasicMergeFields(ref result, data);

            return result;
        }

        /// <summary>
        /// Creates multiple Mail Merges based on the template and data.
        /// </summary>
        /// <param name="template">The template used to insert the data into.</param>
        /// <param name="data">The data to insert into the template. Key for the Merge Field, Value for the content. Calls .ToString() on the value.</param>
        /// <returns>The completed Mail Merge in the same order as the data.</returns>
        public static IList<string> Merge(string template, IList<Dataset> data)
        {
            List<string> results = new List<string>(data.Count);

            foreach(Dataset item in data)
            {
                results.Add(Merge(template, item));
            }

            return results;
        }

        /// <summary>
        /// Replaces basic merge fields with their content and displays/hides associated conditional content
        /// </summary>
        /// <param name="text">The template used to insert the data into.</param>
        /// <param name="data">The data to insert into the template. Key for the Merge Field, Value for the content. Calls .ToString() on the value.</param>
        private static void MergeBasicMergeFields(ref string text, Dataset data)
        {
            foreach(KeyValuePair<string, object> row in data)
            {
                string mergeField = MERGE_FIELD_MARKER_LEFT + row.Key + MERGE_FIELD_MARKER_RIGHT;
                int mergeFieldIndex;
                while ((mergeFieldIndex = text.IndexOf(mergeField)) != -1)
                {

                    bool displayContent = row.Value != null && row.Value.ToString() != "";

                    //replace any conditional content to right if it exists
                    if (mergeFieldIndex + mergeField.Length < text.Length - 2 && text[mergeFieldIndex + mergeField.Length] == MERGE_FIELD_MARKER_LEFT && text[mergeFieldIndex + mergeField.Length + 1] == MERGE_FIELD_MARKER_CONDITIONAL)
                    {
                        DisplayConditionalContentRight(ref text, mergeFieldIndex + mergeField.Length, displayContent);
                    }

                    //replace the merge field tag itself
                    if (displayContent)
                    {
                        DisplayMergefieldContent(ref text, mergeFieldIndex, mergeField, row.Value.ToString());
                    }
                    else
                    {
                        DisplayMergefieldContent(ref text, mergeFieldIndex, mergeField, string.Empty);
                    }

                    //display any conditional content to left if it exists
                    if (mergeFieldIndex >= 2 && text[mergeFieldIndex - 1] == MERGE_FIELD_MARKER_RIGHT && text[mergeFieldIndex - 2] == MERGE_FIELD_MARKER_CONDITIONAL)
                    {
                        DisplayConditionalContentLeft(ref text, mergeFieldIndex - 1, displayContent);
                    }
                }
            }
        }

        private static void MergeIfMergeFields(ref string text, Dataset data)
        {

        }

        /// <summary>
        /// Display or hide merge field content
        /// </summary>
        /// <param name="text"></param>
        /// <param name="mergeFieldIndex"></param>
        /// <param name="mergeField"></param>
        /// <param name="content"></param>
        private static void DisplayMergefieldContent(ref string text, int mergeFieldIndex, string mergeField, string content)
        {
            text = text.Substring(0, mergeFieldIndex) + content + text.Substring(mergeFieldIndex + mergeField.Length);
        }

        /// <summary>
        /// Display or hide conditional content to the left of the starting point
        /// </summary>
        /// <param name="text">The text to be altered</param>
        /// <param name="start">The start position for the search (right-most merge field marker)</param>
        /// <param name="displayContent">Display or hide the merge field conditional content</param>
        private static void DisplayConditionalContentLeft(ref string text, int start, bool displayContent)
        {
            int level = 0;
            int pos = start - 1;
            for(; pos > 0; pos--)
            {
                if (text[pos] == MERGE_FIELD_MARKER_RIGHT)
                {
                    level++;
                }
                else if (text[pos] == MERGE_FIELD_MARKER_LEFT)
                {
                    if(level == 0)
                    {
                        break;
                    }
                    else
                    {
                        level--;
                    }
                }
            }

            if (displayContent)
            {
                //Combine the text before, between and after the merge field markers
                text = text.Substring(0, pos) + text.Substring(pos + 1, start - pos - 2) + text.Substring(start + 1);
            }
            else
            {
                //Combine the text before and after the merge field markers
                text = text.Substring(0, pos) + text.Substring(start + 1);
            }
        }

        /// <summary>
        /// Display or hide conditional content to the right of the starting point
        /// </summary>
        /// <param name="text">The text to be altered</param>
        /// <param name="start">The start position for the search (left-most merge field marker)</param>
        /// <param name="displayContent">Display or hide the merge field conditional content</param>
        private static void DisplayConditionalContentRight(ref string text, int start, bool displayContent)
        {
            int level = 0;
            int pos = start + 1;
            for(; pos < text.Length - 1; pos++)
            {
                if(text[pos] == MERGE_FIELD_MARKER_LEFT)
                {
                    level++;
                }
                else if(text[pos] == MERGE_FIELD_MARKER_RIGHT)
                {
                    if(level == 0)
                    {
                        break;
                    }
                    else
                    {
                        level--;
                    }
                }
            }

            if (displayContent)
            {
                //Combine the text before, between and after the merge field markers
                text = text.Substring(0, start) + text.Substring(start + 2, pos - start - 2) + text.Substring(pos + 1);
            }
            else
            {
                //Combine the text before and after the merge field markers
                text = text.Substring(0, start) + text.Substring(pos + 1);
            }
        }
    }
}

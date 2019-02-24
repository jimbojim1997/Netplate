using System;
using System.Collections.Generic;

using Dataset = System.Collections.Generic.IDictionary<string, object>;

namespace Netplate
{
    public static class MailMerge
    {
        private const char MERGE_FIELD_MARKER_LEFT = '[';
        private const char MERGE_FIELD_MARKER_RIGHT = ']';
        private const char MERGE_FIELD_MARKER_CONDITIONAL = '~';
        private const char MERGE_FIELD_MARKER_STRING = '"';
        private const char MERGE_FIELD_MARKER_ESCAPE = '\\';
        private const char MERGE_FIELD_MARKER_NUMBER_END = ' ';
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
            string mergeFieldMarker = MERGE_FIELD_MARKER_LEFT + MERGE_FIELD_MARKER_IF;
            int mergeFieldIndex;
            while ((mergeFieldIndex = text.IndexOf(mergeFieldMarker)) != -1)
            {
                int searchPosition = mergeFieldIndex + MERGE_FIELD_MARKER_IF.Length + 1;
                bool isValidMergeField = false;
                string mergeField;
                object leftOperand = null;
                object rightOperand = null;
                string op = null;

                mergeField = GetMergeFieldFromPosition(ref text, mergeFieldIndex, out _);
                isValidMergeField = mergeField != null;

                //parse left operand
                if (isValidMergeField)
                {
                    searchPosition++;
                    leftOperand =
                        (object) GetMergeFieldFromPosition(ref text, searchPosition, out searchPosition) ??
                        (object) GetStringFromPosition(ref text, searchPosition, out searchPosition) ??
                        (object) GetNumberFromPosition(ref text, searchPosition, out searchPosition);

                    isValidMergeField = leftOperand != null;
                }

                //parse operator
                if (isValidMergeField)
                {
                    searchPosition++;
                    op = GetOperandFromPosition(ref text, searchPosition, out searchPosition);

                    isValidMergeField = op != null;
                }

                //parse right operand
                if (isValidMergeField)
                {
                    searchPosition++;
                    rightOperand =
                        (object)GetMergeFieldFromPosition(ref text, searchPosition, out searchPosition) ??
                        (object)GetStringFromPosition(ref text, searchPosition, out searchPosition) ??
                        (object)GetNumberFromPosition(ref text, searchPosition, out searchPosition);

                    isValidMergeField = rightOperand != null;
                }

                //do the final merge
                bool displayContent = false;
                if (isValidMergeField)
                {
                    if(leftOperand.GetType() == typeof(string))
                    {
                        leftOperand = Merge(leftOperand.ToString(), data);
                        string tmps = leftOperand.ToString();
                        object tmpo = GetNumberFromPosition(ref tmps, 0, out _);
                        if (tmpo != null) leftOperand = tmpo;
                    }

                    if (rightOperand.GetType() == typeof(string))
                    {
                        rightOperand = Merge(leftOperand.ToString(), data);
                        string tmps = rightOperand.ToString();
                        object tmpo = GetNumberFromPosition(ref tmps, 0, out _);
                        if (tmpo != null) rightOperand = tmpo;
                    }

                    //for both string and number
                    switch (op)
                    {
                        case MERGE_FIELD_IF_OPERATOR_EQUAL:
                            displayContent = leftOperand.ToString().Equals(rightOperand.ToString());
                            break;
                        case MERGE_FIELD_IF_OPERATOR_NOT_EQUAL:
                            displayContent = !leftOperand.ToString().Equals(rightOperand.ToString());
                            break;
                    }

                    //for number (decimal)
                    if(leftOperand.GetType() == typeof(decimal) && rightOperand.GetType() == typeof(decimal))
                    {
                        switch (op)
                        {
                            case MERGE_FIELD_IF_OPERATOR_GREATER_THAN:
                                displayContent = (decimal)leftOperand > (decimal)rightOperand;
                                break;
                            case MERGE_FIELD_IF_OPERATOR_LESS_THAN:
                                displayContent = (decimal)leftOperand < (decimal)rightOperand;
                                break;
                            case MERGE_FIELD_IF_OPERATOR_GREATER_OR_EQUAL:
                                displayContent = (decimal)leftOperand >= (decimal)rightOperand;
                                break;
                            case MERGE_FIELD_IF_OPERATOR_LESS_OR_EQUAL:
                                displayContent = (decimal)leftOperand <= (decimal)rightOperand;
                                break;
                        }
                    }

                }

                //replace any conditional content to right if it exists
                if (mergeFieldIndex + mergeField.Length < text.Length - 2 && text[mergeFieldIndex + mergeField.Length] == MERGE_FIELD_MARKER_LEFT && text[mergeFieldIndex + mergeField.Length + 1] == MERGE_FIELD_MARKER_CONDITIONAL)
                {
                    DisplayConditionalContentRight(ref text, mergeFieldIndex + mergeField.Length, displayContent);
                }

                //replace the merge field tag itself
                DisplayMergefieldContent(ref text, mergeFieldIndex, mergeField, string.Empty);

                //display any conditional content to left if it exists
                if (mergeFieldIndex >= 2 && text[mergeFieldIndex - 1] == MERGE_FIELD_MARKER_RIGHT && text[mergeFieldIndex - 2] == MERGE_FIELD_MARKER_CONDITIONAL)
                {
                    DisplayConditionalContentLeft(ref text, mergeFieldIndex - 1, displayContent);
                }
            }
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

        /// <summary>
        /// Get the merge field from the start position
        /// </summary>
        /// <param name="text">Text to search in</param>
        /// <param name="start">Start position of the search (left-most merge field marker)</param>
        /// <param name="end">The end index of the merge field</param>
        /// <returns>The merge field or null on failure</returns>
        private static string GetMergeFieldFromPosition(ref string text, int start, out int end)
        {
            end = start;
            if (text[start] != MERGE_FIELD_MARKER_LEFT) return null;

            int level = 0;
            int pos = start + 1;

            for (; pos < text.Length - 1; pos++)
            {
                if (text[pos] == MERGE_FIELD_MARKER_LEFT)
                {
                    level++;
                }
                else if (text[pos] == MERGE_FIELD_MARKER_RIGHT)
                {
                    if (level == 0)
                    {
                        break;
                    }
                    else
                    {
                        level--;
                    }
                }
            }

            int length = pos - start + 1;
            if(length < text.Length - 1)
            {
                end = pos;
                return text.Substring(start, pos - start + 1);
            }

            return null;
        }

        /// <summary>
        /// Gets a string from the start position
        /// </summary>
        /// <param name="text">Text to search in</param>
        /// <param name="start">Start position of the search (left-most string marker)</param>
        /// <param name="end">The end index of the string</param>
        /// <returns></returns>
        private static string GetStringFromPosition(ref string text, int start, out int end)
        {
            end = start;
            return null;
        }

        /// <summary>
        /// Gets a number from the start position
        /// </summary>
        /// <param name="text">Text to search in</param>
        /// <param name="start">Start position of the search (left-most digit)</param>
        /// <param name="end">The end index of the number</param>
        /// <returns>The number or null on failure</returns>
        private static Nullable<decimal> GetNumberFromPosition(ref string text, int start, out int end)
        {
            decimal result = default(decimal);

            int numberEndPos = text.IndexOf(MERGE_FIELD_MARKER_NUMBER_END, start);
            if (numberEndPos == -1) numberEndPos = text.IndexOf(MERGE_FIELD_MARKER_RIGHT, start);
            if (numberEndPos == -1) numberEndPos = text.Length;
            if (numberEndPos == -1)
            {
                end = start;
                return null;
            }

            if (decimal.TryParse(text.Substring(start, numberEndPos - start), out result))
            {
                end = numberEndPos;
                return result;
            }
            else
            {
                end = start;
                return null;
            }
        }

        /// <summary>
        /// Gets the first operand from the current position
        /// </summary>
        /// <param name="text">The text to search</param>
        /// <param name="start">The start index of the search</param>
        /// <param name="end">The end position of the search</param>
        /// <returns>The operand</returns>
        private static string GetOperandFromPosition(ref string text, int start, out int end)
        {
            int operandPos = 0;
            string operand = GetLeftMostInstanceOf(ref text, new string[]{
                MERGE_FIELD_IF_OPERATOR_EQUAL,
                MERGE_FIELD_IF_OPERATOR_NOT_EQUAL,
                MERGE_FIELD_IF_OPERATOR_GREATER_THAN,
                MERGE_FIELD_IF_OPERATOR_LESS_THAN,
                MERGE_FIELD_IF_OPERATOR_GREATER_OR_EQUAL,
                MERGE_FIELD_IF_OPERATOR_LESS_OR_EQUAL
            }, out operandPos);

            if (operand != null)
            {
                end = operandPos + operand.Length;
                return operand;
            }

            end = start;
            return null;
        }

        /// <summary>
        /// Gets the left-most item of items from the text
        /// </summary>
        /// <param name="text">The text to search</param>
        /// <param name="items">The items to find</param>
        /// <param name=index">The index of the first item</param>
        /// <returns></returns>
        private static string GetLeftMostInstanceOf(ref string text, string[] items, out int index)
        {
            string result = null;
            index = int.MaxValue;

            foreach(string item in items)
            {
                int i = text.IndexOf(item);
                if(i <= index && i != -1)
                {
                    result = item;
                    index = i;
                }
            }

            return result;
        }
    }
}

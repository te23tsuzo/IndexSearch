/*
 * **********************************
 * Windows Searchを使ったインデックス検索
 * パスと用語を指定するとWindows Searchを使って検索します。
 * 作成者　旭哲男
 * 作成日　2021/07/24
 
 * Copyright 2021 Tetsuo Asahi
 * Permission is hereby granted, free of charge, to any person obtaining a copy of this software and 
 * associated documentation files (the "Software"), to deal in the Software without restriction, 
 * including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, 
 * and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, 
 * subject to the following conditions:
 * 
 * The above copyright notice and this permission notice shall be included in all copies or substantial
 * portions of the Software.
 * 
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, 
 * INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. 
 * IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, 
 * WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR 
 * THE USE OR OTHER DEALINGS IN THE SOFTWARE.

 */

using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Text;

namespace IndexSearch
{
    class Program
    {
        /// <summary>
        /// Index Searchの本体
        /// </summary>
        /// <param name="args">
        /// 　--path 検索対象のフルパス。WindowsのIndexが設定されていないパスを指定すると、検索結果が出ません。
        /// 　検索結果が出てこない場合は、”Windows Searchの設定”を確認してください。
        /// 　指定しない場合は、実効ディレクトリは以下を検索します。
        /// 　--word 検索対象の用語。""でくくってください。"Azure SQL"のように連続した単語も指定できます。
        /// 　Windows Search SQLのCONTAINS術後の引数となります。以下も参照してください。
        /// 　https://docs.microsoft.com/ja-jp/windows/win32/search/-search-sql-contains
        /// </param>
        static void Main(string[] args)
        {

            List<KeyValuePair<string, string>> argls = makeParams(args);

            string query = GetQuery(argls);

            using (OleDbConnection objConnection =
                new OleDbConnection
                ("Provider=Search.CollatorDSO.1;Extended?Properties='Application=Windows';"))
            {
                objConnection.Open();
                OleDbCommand cmd = new OleDbCommand(query, objConnection);
                using (OleDbDataReader rdr = cmd.ExecuteReader())
                {
                    for (int i = 0; i < rdr.FieldCount; i++)
                    {
                        Console.Write(rdr.GetName(i));
                        Console.Write('\t');
                    }
                    while (rdr.Read())
                    {
                        Console.WriteLine();
                        for (int i = 0; i < rdr.FieldCount; i++)
                        {
                            Console.Write(rdr[i]);
                            Console.Write('\t');
                        }
                    }
                    //Console.ReadKey();
                }
            }

        }

        private static string GetQuery(List<KeyValuePair<string, string>> argls)
        {
            string path;
            string word;


            try
            {
                path = argls.Find(x => x.Key.StartsWith("--path")).Value.Replace('\\', '/');
            } catch (System.NullReferenceException)
            {
                path = System.Environment.CurrentDirectory.Replace('\\', '/');
            }
            
            try
            {
                word = argls.Find(y => y.Key.StartsWith("--word")).Value;
            } catch (System.NullReferenceException)
            {
                word = "";
                System.Console.Error.WriteLine("--word キーワードを指定してください。");
                System.Environment.Exit(1);
            }
            

            StringBuilder sql = new StringBuilder("SELECT TOP 1000 System.ItemPathDisplay FROM SYSTEMINDEX");
            sql.Append(" WHERE CONTAINS('\"{1}\"') AND scope='file://{0}'");
            sql.Append(" ORDER BY System.ItemUrl");

            return string.Format(sql.ToString(), path, word);
        }

        public static List<KeyValuePair<string, string>> makeParams(string[] args)
        {
            List<KeyValuePair<string, string>> ls = new List<KeyValuePair<string, string>>();

            for (int i=0;i<args.Length;i++)
            {
                if (args[i].StartsWith("--"))
                {
                    string _key = args[i];
                    string _value = args[++i];

                    if ( _value.StartsWith("'") || _value.StartsWith("\"")) {
                        
                        for (string _value2 = args[++i]; 
                            (!_value2.StartsWith("--") && i<args.Length);
                            _value2 = args[++i]
                            )
                        {
                            _value += " " + _value2;
                            if (_value2.EndsWith("'") || _value2.EndsWith("\"")) break;
                        }
                        _value = _value.Replace("'", "");
                    } 
                        
                    KeyValuePair<string, string> kv = new KeyValuePair<string, string>(_key, _value );
                    ls.Add(kv);
                }
            }
            return ls;
        }

    }
}

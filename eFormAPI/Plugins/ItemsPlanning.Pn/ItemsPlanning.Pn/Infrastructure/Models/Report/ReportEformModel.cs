﻿/*
The MIT License (MIT)

Copyright (c) 2007 - 2021 Microting A/S

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
*/

namespace ItemsPlanning.Pn.Infrastructure.Models.Report
{
    using System.Collections.Generic;

    public class ReportEformModel
    {
        public string TemplateName { get; set; }

        public string FromDate { get; set; }

        public string ToDate { get; set; }

        public string TableName { get; set; }

        public List<string> DescriptionBlocks { get; set; }
            = new List<string>();

        public List<KeyValuePair<List<string>, List<string>>> ImageNames { get; set; }
            = new List<KeyValuePair<List<string>, List<string>>>();

        public List<KeyValuePair<int, string>> ItemHeaders { get; set; }
            = new List<KeyValuePair<int, string>>();

        public List<ReportEformItemModel> Items { get; set; }
            = new List<ReportEformItemModel>();

        public List<ReportEformPostModel> Posts { get; set; }
            = new List<ReportEformPostModel>();
        
        public ReportEformTextHeaderModel TextHeaders { get; set; }

        public List<string> NameTagsInEndPage { get; set; }
            = new List<string>();
    }
}

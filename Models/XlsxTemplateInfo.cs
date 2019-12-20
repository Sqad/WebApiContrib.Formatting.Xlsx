﻿namespace WebApiContrib.Formatting.Xlsx.Models
{
    public class XlsxTemplateInfo
    {
        public XlsxTemplateInfo(string path, string password)
        {
            Path = path;
            Password = password;
        }

        public string Path { get; }
        public string Password { get; }
    }
}
using System;
using System.Linq;
using System.Xml;

namespace ProjectX.Dict
{
    public class GenId
    {
        private int NumPlace { get; set; }
        private int NumIndex { get; set; }
        private char CharIndex { get; set; }

        public GenId(char charIndex, int numIndex, int numPlace)
        {
            CharIndex = charIndex;
            NumIndex = numIndex;
            NumPlace = numPlace >= NumP(numIndex) ? numPlace : NumP(numIndex);
        }

        public string ThisVal()
        {
            return CharIndex + String.Concat(Enumerable.Repeat("0", NumPlace - NumP(NumIndex))) + NumIndex;
        }

        public string NexVal()
        {
            if (NumP(++NumIndex) > NumPlace)
            {
                if (CharIndex == 'Z')
                {
                    CharIndex = 'A';
                    NumPlace++;
                }
                else
                {
                    CharIndex++;
                }
                NumIndex = 0;
            }
            return ThisVal();
        }

        private static int NumP(int n)
        {
            int c = 0;
            do
            {
                c++;
            }
            while ((n = n / 10) != 0);
            return c;
        }

        public XmlElement GetXmlNode(XmlDocument document)
        {
            XmlElement element = document.CreateElement("settings");
            XmlElement e = document.CreateElement("charIndex");
            e.InnerText = CharIndex.ToString();
            element.AppendChild(e);
            e = document.CreateElement("numIndex");
            e.InnerText = NumIndex.ToString();
            element.AppendChild(e);
            e = document.CreateElement("numPlace");
            e.InnerText = NumPlace.ToString();
            element.AppendChild(e);
            return element;
        }

        public XmlElement GetXmlNode(XmlDocument document,string nameNode)
        {
            XmlElement element = document.CreateElement(nameNode);
            XmlElement e = document.CreateElement("charIndex");
            e.InnerText = CharIndex.ToString();
            element.AppendChild(e);
            e = document.CreateElement("numIndex");
            e.InnerText = NumIndex.ToString();
            element.AppendChild(e);
            e = document.CreateElement("numPlace");
            e.InnerText = NumPlace.ToString();
            element.AppendChild(e);
            return element;
        }

    }
}
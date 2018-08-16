using CSSPEnums;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Principal;
using System.Text;
using System.Threading.Tasks;
using CSSPTaskRunnerHelper.Resources;

namespace CSSPTaskRunnerHelper
{
    public partial class TaskRunnerHelperBase
    {
        #region Variables
        #endregion Variables

        #region Properties
        public List<LanguageEnum> LanguageListAllowable = new List<LanguageEnum>() { LanguageEnum.en, LanguageEnum.fr };
        #endregion Properties

        #region Constructors
        public TaskRunnerHelperBase()
        {
        }
        #endregion Constructors

        #region Events
        public virtual void MessageEvent(MessageEventArgs e)
        {
            MessageHandler?.Invoke(this, e);
        }
        public event EventHandler<MessageEventArgs> MessageHandler;
        #endregion Events

        #region Functions public
        #endregion Functions public

        #region Functions protected
        protected List<TextLanguage> GetTextLanguageList()
        {
            List<TextLanguage> TextLanguageList = new List<TextLanguage>();
            foreach (LanguageEnum lang in LanguageListAllowable)
            {
                TextLanguageList.Add(new TextLanguage() { Language = lang, Text = "" });
            }
            return TextLanguageList;
        }
        protected List<TextLanguage> GetTextLanguageList(string ResString)
        {
            List<TextLanguage> TextLanguageList = new List<TextLanguage>();
            foreach (LanguageEnum lang in LanguageListAllowable)
            {
                string TempText = CSSPTaskRunnerHelperRes.ResourceManager.GetString(ResString, new System.Globalization.CultureInfo(lang.ToString()));
                TextLanguageList.Add(new TextLanguage() { Language = lang, Text = TempText });
            }
            return TextLanguageList;
        }
        protected List<TextLanguage> GetTextLanguageFormat1List(string ResString, string Param1)
        {
            List<TextLanguage> TextLanguageList = new List<TextLanguage>();
            foreach (LanguageEnum lang in LanguageListAllowable)
            {
                string TempText = string.Format(CSSPTaskRunnerHelperRes.ResourceManager.GetString(ResString, new System.Globalization.CultureInfo(lang.ToString())), Param1);
                TextLanguageList.Add(new TextLanguage() { Language = lang, Text = TempText });
            }
            return TextLanguageList;
        }
        public List<TextLanguage> GetTextLanguageFormat2List(string ResString, string Param1, string Param2)
        {
            List<TextLanguage> TextLanguageList = new List<TextLanguage>();
            foreach (LanguageEnum lang in LanguageListAllowable)
            {
                string TempText = string.Format(CSSPTaskRunnerHelperRes.ResourceManager.GetString(ResString, new System.Globalization.CultureInfo(lang.ToString())), Param1, Param2);
                TextLanguageList.Add(new TextLanguage() { Language = lang, Text = TempText });
            }
            return TextLanguageList;
        }
        public List<TextLanguage> GetTextLanguageFormat3List(string ResString, string Param1, string Param2, string Param3)
        {
            List<TextLanguage> TextLanguageList = new List<TextLanguage>();
            foreach (LanguageEnum lang in LanguageListAllowable)
            {
                string TempText = string.Format(CSSPTaskRunnerHelperRes.ResourceManager.GetString(ResString, new System.Globalization.CultureInfo(lang.ToString())), Param1, Param2, Param3);
                TextLanguageList.Add(new TextLanguage() { Language = lang, Text = TempText });
            }
            return TextLanguageList;
        }
        public List<TextLanguage> GetTextLanguageFormat4List(string ResString, string Param1, string Param2, string Param3, string Param4)
        {
            List<TextLanguage> TextLanguageList = new List<TextLanguage>();
            foreach (LanguageEnum lang in LanguageListAllowable)
            {
                string TempText = string.Format(CSSPTaskRunnerHelperRes.ResourceManager.GetString(ResString, new System.Globalization.CultureInfo(lang.ToString())), Param1, Param2, Param3, Param4);
                TextLanguageList.Add(new TextLanguage() { Language = lang, Text = TempText });
            }
            return TextLanguageList;
        }
        public List<TextLanguage> GetTextLanguageFormat5List(string ResString, string Param1, string Param2, string Param3, string Param4, string Param5)
        {
            List<TextLanguage> TextLanguageList = new List<TextLanguage>();
            foreach (LanguageEnum lang in LanguageListAllowable)
            {
                string TempText = string.Format(CSSPTaskRunnerHelperRes.ResourceManager.GetString(ResString, new System.Globalization.CultureInfo(lang.ToString())), Param1, Param2, Param3, Param4, Param5);
                TextLanguageList.Add(new TextLanguage() { Language = lang, Text = TempText });
            }
            return TextLanguageList;
        }
        #endregion Functions protected

        #region Functions private
        #endregion Functions private

        #region Sub Classes
        public class MessageEventArgs : EventArgs
        {
            public MessageEventArgs(string Message, MessageTypeEnum Type)
            {
                this.Message = Message;
                this.Type = MessageTypeEnum.Status;
            }

            public string Message { get; set; }
            public MessageTypeEnum Type { get; set; }
        }
        public class TextLanguage
        {
            public LanguageEnum Language { get; set; }
            public string Text { get; set; }
        }
        #endregion Sub Classes

        #region Enums
        public enum MessageTypeEnum
        {
            Error = 1,
            Status = 2,
            Permanent = 3,
            Clear = 4,
        }
        #endregion Enums
    }
}

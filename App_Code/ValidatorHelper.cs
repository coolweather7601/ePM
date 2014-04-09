using System;
using System.Data;
using System.Configuration;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using System.Text.RegularExpressions;

namespace EPM_Web.Alan.Common
{
    /// <summary>
    /// ValidatorHelper 的摘要描述
    /// </summary>
    public class ValidatorHelper
    {
        public struct RegularExp
        {
            public const string Chinese = @"^[\u4E00-\u9FA5\uF900-\uFA2D]+$";
            public const string Color = "^#[a-fA-F0-9]{6}";
            public const string Date = @"^((((1[6-9]|[2-9]\d)\d{2})-(0?[13578]|1[02])-(0?[1-9]|[12]\d|3[01]))|(((1[6-9]|[2-9]\d)\d{2})-(0?[13456789]|1[012])-(0?[1-9]|[12]\d|30))|(((1[6-9]|[2-9]\d)\d{2})-0?2-(0?[1-9]|1\d|2[0-8]))|(((1[6-9]|[2-9]\d)(0[48]|[2468][048]|[13579][26])|((16|[2468][048]|[3579][26])00))-0?2-29-))$";
            public const string DateTime = @"^((((1[6-9]|[2-9]\d)\d{2})-(0?[13578]|1[02])-(0?[1-9]|[12]\d|3[01]))|(((1[6-9]|[2-9]\d)\d{2})-(0?[13456789]|1[012])-(0?[1-9]|[12]\d|30))|(((1[6-9]|[2-9]\d)\d{2})-0?2-(0?[1-9]|1\d|2[0-8]))|(((1[6-9]|[2-9]\d)(0[48]|[2468][048]|[13579][26])|((16|[2468][048]|[3579][26])00))-0?2-29-)) (20|21|22|23|[0-1]?\d):[0-5]?\d:[0-5]?\d$";
            public const string Email = @"^[\w-]+(\.[\w-]+)*@[\w-]+(\.[\w-]+)+$";
            public const string Float = @"^(-?\d+)(\.\d+)?$";
            public const string ImageFormat = @"\.(?i:jpg|bmp|gif|ico|pcx|jpeg|tif|png|raw|tga)$";
            public const string Integer = @"^-?\d+$";
            public const string IP = @"^(\d{1,2}|1\d\d|2[0-4]\d|25[0-5])\.(\d{1,2}|1\d\d|2[0-4]\d|25[0-5])\.(\d{1,2}|1\d\d|2[0-4]\d|25[0-5])\.(\d{1,2}|1\d\d|2[0-4]\d|25[0-5])$";
            public const string Letter = "^[A-Za-z]+$";
            public const string LowerLetter = "^[a-z]+$";
            public const string MinusFloat = @"^(-(([0-9]+\.[0-9]*[1-9][0-9]*)|([0-9]*[1-9][0-9]*\.[0-9]+)|([0-9]*[1-9][0-9]*)))$";
            public const string MinusInteger = "^-[0-9]*[1-9][0-9]*$";
            public const string Mobile = "^0{0,1}13[0-9]{9}$";
            public const string NumbericOrLetterOrChinese = @"^[A-Za-z0-9\u4E00-\u9FA5\uF900-\uFA2D]+$";
            public const string Numeric = "^[0-9]+$";
            public const string NumericOrLetter = "^[A-Za-z0-9]+$";
            public const string NumericOrLetterOrUnderline = @"^\w+$";
            public const string PlusFloat = @"^(([0-9]+\.[0-9]*[1-9][0-9]*)|([0-9]*[1-9][0-9]*\.[0-9]+)|([0-9]*[1-9][0-9]*))$";
            public const string PlusInteger = "^[0-9]*[1-9][0-9]*$";
            public const string Telephone = @"(\d+-)?(\d{4}-?\d{7}|\d{3}-?\d{8}|^\d{7,8})(-\d+)?";
            public const string UnMinusFloat = @"^\d+(\.\d+)?$";
            public const string UnMinusInteger = @"\d+$";
            public const string UnPlusFloat = @"^((-\d+(\.\d+)?)|(0+(\.0+)?))$";
            public const string UnPlusInteger = @"^((-\d+)|(0+))$";
            public const string UpperLetter = "^[A-Z]+$";
            public const string Url = @"^http(s)?://([\w-]+\.)+[\w-]+(/[\w- ./?%&=]*)?$";
        }

        /// <summary>
        /// 判斷字串是否與指定正則運算式匹配
        /// </summary>
        /// <param name="input">要驗證的字串</param>
        /// <param name="regularExp">正則運算式</param>
        /// <returns>驗證通過返回true</returns>
        public bool IsMatch(string input, string regularExp)
        {
            return Regex.IsMatch(input, regularExp);
        }

        /// <summary>
        /// 驗證非負整數（正整數 + 0）
        /// </summary>
        /// <param name="input">要驗證的字串</param>
        /// <returns>驗證通過返回true</returns>
        public bool IsUnMinusInt(string input)
        {
            return Regex.IsMatch(input, RegularExp.UnMinusInteger);
        }

        /// <summary>
        /// 驗證正整數
        /// </summary>
        /// <param name="input">要驗證的字串</param>
        /// <returns>驗證通過返回true</returns>
        public bool IsPlusInt(string input)
        {
            return Regex.IsMatch(input, RegularExp.PlusInteger);
        }

        /// <summary>
        /// 驗證非正整數（負整數 + 0） 
        /// </summary>
        /// <param name="input">要驗證的字串</param>
        /// <returns>驗證通過返回true</returns>
        public bool IsUnPlusInt(string input)
        {
            return Regex.IsMatch(input, RegularExp.UnPlusInteger);
        }

        /// <summary>
        /// 驗證負整數
        /// </summary>
        /// <param name="input">要驗證的字串</param>
        /// <returns>驗證通過返回true</returns>
        public bool IsMinusInt(string input)
        {
            return Regex.IsMatch(input, RegularExp.MinusInteger);
        }

        /// <summary>
        /// 驗證整數
        /// </summary>
        /// <param name="input">要驗證的字串</param>
        /// <returns>驗證通過返回true</returns>
        public bool IsInt(string input)
        {
            return Regex.IsMatch(input, RegularExp.Integer);
        }

        /// <summary>
        /// 驗證非負浮點數（正浮點數 + 0）
        /// </summary>
        /// <param name="input">要驗證的字串</param>
        /// <returns>驗證通過返回true</returns>
        public bool IsUnMinusFloat(string input)
        {
            return Regex.IsMatch(input, RegularExp.UnMinusFloat);
        }

        /// <summary>
        /// 驗證正浮點數
        /// </summary>
        /// <param name="input">要驗證的字串</param>
        /// <returns>驗證通過返回true</returns>
        public bool IsPlusFloat(string input)
        {
            return Regex.IsMatch(input, RegularExp.PlusFloat);
        }

        /// <summary>
        /// 驗證非正浮點數（負浮點數 + 0）
        /// </summary>
        /// <param name="input">要驗證的字串</param>
        /// <returns>驗證通過返回true</returns>
        public bool IsUnPlusFloat(string input)
        {
            return Regex.IsMatch(input, RegularExp.UnPlusFloat);
        }

        /// <summary>
        /// 驗證負浮點數
        /// </summary>
        /// <param name="input">要驗證的字串</param>
        /// <returns>驗證通過返回true</returns>
        public bool IsMinusFloat(string input)
        {
            return Regex.IsMatch(input, RegularExp.MinusFloat);
        }

        /// <summary>
        /// 驗證浮點數
        /// </summary>
        /// <param name="input">要驗證的字串</param>
        /// <returns>驗證通過返回true</returns>
        public bool IsFloat(string input)
        {
            if (input.Equals("N/A")) { return true; }
            return Regex.IsMatch(input, RegularExp.Float);
        }

        /// <summary>
        /// 驗證由26個英文字母組成的字串
        /// </summary>
        /// <param name="input">要驗證的字串</param>
        /// <returns>驗證通過返回true</returns>
        public bool IsLetter(string input)
        {
            return Regex.IsMatch(input, RegularExp.Letter);
        }

        /// <summary>
        /// 驗證由中文組成的字串
        /// </summary>
        /// <param name="input">要驗證的字串</param>
        /// <returns>驗證通過返回true</returns>
        public bool IsChinese(string input)
        {
            return Regex.IsMatch(input, RegularExp.Chinese);
        }

        /// <summary>
        /// 驗證由26個英文字母的大寫組成的字串
        /// </summary>
        /// <param name="input">要驗證的字串</param>
        /// <returns>驗證通過返回true</returns>
        public bool IsUpperLetter(string input)
        {
            return Regex.IsMatch(input, RegularExp.UpperLetter);
        }

        /// <summary>
        /// 驗證由26個英文字母的小寫組成的字串
        /// </summary>
        /// <param name="input">要驗證的字串</param>
        /// <returns>驗證通過返回true</returns>
        public bool IsLowerLetter(string input)
        {
            return Regex.IsMatch(input, RegularExp.LowerLetter);
        }

        /// <summary>
        /// 驗證由數位和26個英文字母組成的字串
        /// </summary>
        /// <param name="input">要驗證的字串</param>
        /// <returns>驗證通過返回true</returns>
        public bool IsNumericOrLetter(string input)
        {
            return Regex.IsMatch(input, RegularExp.NumericOrLetter);
        }

        /// <summary>
        /// 驗證由數位組成的字串
        /// </summary>
        /// <param name="input">要驗證的字串</param>
        /// <returns>驗證通過返回true</returns>
        public bool IsNumeric(string input)
        {
            return Regex.IsMatch(input, RegularExp.Numeric);
        }
        /// <summary>
        /// 驗證由數位和26個英文字母或中文組成的字串
        /// </summary>
        /// <param name="input">要驗證的字串</param>
        /// <returns>驗證通過返回true</returns>
        public bool IsNumericOrLetterOrChinese(string input)
        {
            return Regex.IsMatch(input, RegularExp.NumbericOrLetterOrChinese);
        }

        /// <summary>
        /// 驗證由數位、26個英文字母或者下劃線組成的字串
        /// </summary>
        /// <param name="input">要驗證的字串</param>
        /// <returns>驗證通過返回true</returns>
        public bool IsNumericOrLetterOrUnderline(string input)
        {
            return Regex.IsMatch(input, RegularExp.NumericOrLetterOrUnderline);
        }

        /// <summary>
        /// 驗證email地址
        /// </summary>
        /// <param name="input">要驗證的字串</param>
        /// <returns>驗證通過返回true</returns>
        public bool IsEmail(string input)
        {
            return Regex.IsMatch(input, RegularExp.Email);
        }

        /// <summary>
        /// 驗證URL
        /// </summary>
        /// <param name="input">要驗證的字串</param>
        /// <returns>驗證通過返回true</returns>
        public bool IsUrl(string input)
        {
            return Regex.IsMatch(input, RegularExp.Url);
        }

        /// <summary>
        /// 驗證電話號碼
        /// </summary>
        /// <param name="input">要驗證的字串</param>
        /// <returns>驗證通過返回true</returns>
        public bool IsTelephone(string input)
        {
            return Regex.IsMatch(input, RegularExp.Telephone);
        }

        /// <summary>
        /// 驗證手機號碼
        /// </summary>
        /// <param name="input">要驗證的字串</param>
        /// <returns>驗證通過返回true</returns>
        public bool IsMobile(string input)
        {
            return Regex.IsMatch(input, RegularExp.Mobile);
        }

        /// <summary>
        /// 通過檔副檔名驗證圖像格式
        /// </summary>
        /// <param name="input">要驗證的字串</param>
        /// <returns>驗證通過返回true</returns>
        public bool IsImageFormat(string input)
        {
            return Regex.IsMatch(input, RegularExp.ImageFormat);
        }

        /// <summary>
        /// 驗證IP
        /// </summary>
        /// <param name="input">要驗證的字串</param>
        /// <returns>驗證通過返回true</returns>
        public bool IsIP(string input)
        {
            return Regex.IsMatch(input, RegularExp.IP);
        }

        /// <summary>
        /// 驗證日期（YYYY-MM-DD）
        /// </summary>
        /// <param name="input">要驗證的字串</param>
        /// <returns>驗證通過返回true</returns>
        public bool IsDate(string input)
        {
            return Regex.IsMatch(input, RegularExp.Date);
        }

        /// <summary>
        /// 驗證日期和時間（YYYY-MM-DD HH:MM:SS）
        /// </summary>
        /// <param name="input">要驗證的字串</param>
        /// <returns>驗證通過返回true</returns>
        public bool IsDateTime(string input)
        {
            return Regex.IsMatch(input, RegularExp.DateTime);
        }

        /// <summary>
        /// 驗證顏色（#ff0000）
        /// </summary>
        /// <param name="input">要驗證的字串</param>
        /// <returns>驗證通過返回true</returns>
        public bool IsColor(string input)
        {
            return Regex.IsMatch(input, RegularExp.Color);
        }
    }
}
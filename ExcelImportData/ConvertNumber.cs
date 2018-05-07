using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelImportData
{
    /// <summary>
    /// Read Number by NguyenTueVuong: https://github.com/nguyentuevuong/read-number
    /// </summary>
    public static class ConvertNumber
    {
        /// <summary>
        /// Convert number from decimal to Vietnamese readable
        /// </summary>
        /// <param name="_number">number to convert</param>
        /// <returns>string is Vietnamese readable</returns>
        public static string Convert(decimal _number)
        {
            // if  number equal 0
            if (_number == 0)
                return "Không.";

            //is negative or positive
            bool _startwith = _number < 0 ? true : false;

            string _letter = "";
            //absolute number
            _number = Math.Abs(_number);

            //Format number
            string _source = String.Format("{0:0,0}", _number);

            //find dot in number
            string[] _arrsource = _source.Split(',');

            //divide the unit of number
            int _numunit = _arrsource.Length;
            foreach (string _str in _arrsource)
            {
                if (ThreeNumber2Letter(int.Parse(_str)) != "")
                    _letter += String.Format("{0} {1}, ", ThreeNumber2Letter(int.Parse(_str)), NumUnit(_numunit));
                _numunit--;
            }
            _letter = _letter.Substring(0, _letter.Length - 2);

            if (_letter.StartsWith("không trăm"))
                _letter = _letter.Substring(10, _letter.Length - 10).Trim();
            if (_letter.StartsWith("lẻ"))
                _letter = _letter.Substring(2, _letter.Length - 2).Trim();

            _letter = String.Format("{0} {1}", _startwith ? "âm" : "", _letter).Trim().Replace("lăm trăm", "năm trăm");
            return String.Format("{0}{1}.", _letter.Substring(0, 1).ToUpper(), _letter.Substring(1, _letter.Length - 1).Trim());
        }
        /// <summary>
        /// readable unit number
        /// </summary>
        /// <param name="_unit">unit number</param>
        /// <returns>string is readable unit number</returns>
        private static string NumUnit(int _unit)
        {
            if (_unit < 2) return "";

            switch (_unit)
            {
                //readable unit number
                case 2: return "nghìn";
                case 3: return "triệu";
                case 4: return "tỷ";
                default: return String.Format("{0} {1}", NumUnit(_unit - 3), NumUnit(4));
            }
        }

        /// <summary>
        /// readable 3 number in a unit numbe
        /// </summary>
        /// <param name="_number">3 number</param>
        /// <returns>readable 3 number in a unit numbe</returns>
        private static string ThreeNumber2Letter(int _number)
        {
            int _hunit = 0, _tunit = 0, _nunit = 0;

            if (_number > 0 && _number < 10)// Trường hợp _number = [1-9]
                _nunit = _number;
            else if (_number > 9 && _number < 100) // Trường hợp _number = [10-99]
            {
                _tunit = _number / 10;
                _nunit = _number - (_tunit * 10);
            }
            else if (_number > 99 && _number < 1000)// Trường hợp _number = [100-999]
            {
                _hunit = _number / 100;
                _tunit = (_number - (_hunit * 100)) / 10;
                _nunit = _number - (_hunit * 100) - (_tunit * 10);
            }
            else // Trường hợp _number <> [1-999]
                return "";

            string[] _OneNumber2Letter = { "không", "một", "hai", "ba", "bốn", "lăm", "sáu", "bảy", "tám", "chín" };

            switch (_tunit)
            {
                case 0:
                    if (_nunit == 0)
                        return String.Format("{0} trăm", _OneNumber2Letter[_hunit]);
                    else
                    {
                        if (_nunit == 5)
                            return String.Format("{0} trăm lẻ năm", _OneNumber2Letter[_hunit]);
                        else
                            return String.Format("{0} trăm lẻ {1}", _OneNumber2Letter[_hunit], _OneNumber2Letter[_nunit]);
                    }
                case 1:
                    if (_nunit == 0)
                        return String.Format("{0} trăm mười", _OneNumber2Letter[_hunit]);
                    else
                        return String.Format("{0} trăm mười {1}", _OneNumber2Letter[_hunit], _OneNumber2Letter[_nunit]);
                case 5:
                    return String.Format("{0} trăm năm mươi {1}", _OneNumber2Letter[_hunit], _OneNumber2Letter[_nunit]);
                default:
                    if (_nunit == 0)
                        return String.Format("{0} trăm {1} mươi", _OneNumber2Letter[_hunit], _OneNumber2Letter[_tunit]);
                    else if (_nunit == 1)
                        return String.Format("{0} trăm {1} mươi mốt", _OneNumber2Letter[_hunit], _OneNumber2Letter[_tunit]);
                    else if (_nunit == 4)
                        return String.Format("{0} trăm {1} mươi tư", _OneNumber2Letter[_hunit], _OneNumber2Letter[_tunit]);
                    else
                        return String.Format("{0} trăm {1} mươi {2}", _OneNumber2Letter[_hunit], _OneNumber2Letter[_tunit], _OneNumber2Letter[_nunit]);
            }
        }

    }
}

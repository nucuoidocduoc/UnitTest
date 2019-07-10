using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Units
{
    [Regeneration(RegenerationOption.Manual)]
    [Transaction(TransactionMode.Manual)]
    public class Command : IExternalCommand
    {
        private Document _document;
        private Autodesk.Revit.DB.Units _unitsProject;
        private FormatOptions _lengthFormatOptions;
        private FormatOptions _areaFormatOptions;
        private FormatOptions _volumnFormatOptions;
        private FormatOptions _angleFormatOptions;
        private FormatOptions _slopeFormatOptions;
        private FormatOptions _currencyFormatOptions;
        private FormatOptions _massDensityFormatOptions;

        private UnitSymbolType _unitSymbolTypeLength;
        private UnitSymbolType _unitSymbolTypeArea;
        private UnitSymbolType _unitSymbolTypeVolume;
        private UnitSymbolType _unitSymbolTypeAngle;
        private UnitSymbolType _unitSymbolTypeSlope;
        private UnitSymbolType _unitSymbolTypeCurrency;
        private UnitSymbolType _unitSymbolTypeMassDensity;

        private UnitSymbolType _noneSymbol;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document document = commandData.Application.ActiveUIDocument.Document;
            _unitsProject = document.GetUnits();
            _document = document;

            _noneSymbol = UnitSymbolType.UST_NONE;
            _lengthFormatOptions = _unitsProject.GetFormatOptions(UnitType.UT_Length);
            _areaFormatOptions = _unitsProject.GetFormatOptions(UnitType.UT_Area);
            _volumnFormatOptions = _unitsProject.GetFormatOptions(UnitType.UT_Volume);
            _angleFormatOptions = _unitsProject.GetFormatOptions(UnitType.UT_Angle);
            _slopeFormatOptions = _unitsProject.GetFormatOptions(UnitType.UT_Slope);
            _currencyFormatOptions = _unitsProject.GetFormatOptions(UnitType.UT_Currency);
            _massDensityFormatOptions = _unitsProject.GetFormatOptions(UnitType.UT_MassDensity);

            _unitSymbolTypeLength = _lengthFormatOptions.UnitSymbol;
            _unitSymbolTypeArea = _areaFormatOptions.UnitSymbol;
            _unitSymbolTypeVolume = _volumnFormatOptions.UnitSymbol;
            _unitSymbolTypeAngle = _angleFormatOptions.UnitSymbol;
            _unitSymbolTypeSlope = _slopeFormatOptions.UnitSymbol;
            _unitSymbolTypeCurrency = _currencyFormatOptions.UnitSymbol;
            _unitSymbolTypeMassDensity = _massDensityFormatOptions.UnitSymbol;
            using (Transaction t = new Transaction(document, "offsymbol")) {
                t.Start();
                OffSymbol();
                t.Commit();
            }
            return Result.Succeeded;
        }

        public void OnSymbol()
        {
            _lengthFormatOptions.UnitSymbol = _unitSymbolTypeLength;
            _areaFormatOptions.UnitSymbol = _unitSymbolTypeArea;
            _volumnFormatOptions.UnitSymbol = _unitSymbolTypeVolume;
            _angleFormatOptions.UnitSymbol = _unitSymbolTypeAngle;
            _slopeFormatOptions.UnitSymbol = _unitSymbolTypeSlope;
            _currencyFormatOptions.UnitSymbol = _unitSymbolTypeCurrency;
            _massDensityFormatOptions.UnitSymbol = _unitSymbolTypeMassDensity;

            _unitsProject.SetFormatOptions(UnitType.UT_Length, _lengthFormatOptions);
            _unitsProject.SetFormatOptions(UnitType.UT_Area, _areaFormatOptions);
            _unitsProject.SetFormatOptions(UnitType.UT_Volume, _volumnFormatOptions);
            _unitsProject.SetFormatOptions(UnitType.UT_Angle, _angleFormatOptions);
            _unitsProject.SetFormatOptions(UnitType.UT_Slope, _slopeFormatOptions);
            _unitsProject.SetFormatOptions(UnitType.UT_Currency, _currencyFormatOptions);
            _unitsProject.SetFormatOptions(UnitType.UT_MassDensity, _massDensityFormatOptions);
            _document.SetUnits(_unitsProject);
        }

        public void OffSymbol()
        {
            _lengthFormatOptions.UnitSymbol = _noneSymbol;
            _areaFormatOptions.UnitSymbol = _noneSymbol;
            _volumnFormatOptions.UnitSymbol = _noneSymbol;
            _angleFormatOptions.UnitSymbol = _noneSymbol;
            _slopeFormatOptions.UnitSymbol = _noneSymbol;
            _currencyFormatOptions.UnitSymbol = _noneSymbol;
            _massDensityFormatOptions.UnitSymbol = _noneSymbol;

            _unitsProject.SetFormatOptions(UnitType.UT_Length, _lengthFormatOptions);
            _unitsProject.SetFormatOptions(UnitType.UT_Area, _areaFormatOptions);
            _unitsProject.SetFormatOptions(UnitType.UT_Volume, _volumnFormatOptions);
            _unitsProject.SetFormatOptions(UnitType.UT_Angle, _angleFormatOptions);
            _unitsProject.SetFormatOptions(UnitType.UT_Slope, _slopeFormatOptions);
            _unitsProject.SetFormatOptions(UnitType.UT_Currency, _currencyFormatOptions);
            _unitsProject.SetFormatOptions(UnitType.UT_MassDensity, _massDensityFormatOptions);
            _document.SetUnits(_unitsProject);
        }
    }
}
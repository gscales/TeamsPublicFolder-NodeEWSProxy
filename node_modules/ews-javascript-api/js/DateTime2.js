"use strict";
var DateTime_1 = require("./DateTime");
var DateTime = (function () {
    function DateTime() {
    }
    return DateTime;
}());
// Number of 100ns ticks per time unit
DateTime.TicksPerMillisecond = 10000; //long
DateTime.TicksPerSecond = DateTime.TicksPerMillisecond * 1000; //long
DateTime.TicksPerMinute = DateTime.TicksPerSecond * 60; //long
DateTime.TicksPerHour = DateTime.TicksPerMinute * 60; //long
DateTime.TicksPerDay = DateTime.TicksPerHour * 24; //long
// Number of milliseconds per time unit
DateTime.MillisPerSecond = 1000; //int
DateTime.MillisPerMinute = DateTime.MillisPerSecond * 60; //int
DateTime.MillisPerHour = DateTime.MillisPerMinute * 60; //int
DateTime.MillisPerDay = DateTime.MillisPerHour * 24; //int
// Number of days in a non-leap year
DateTime.DaysPerYear = 365; //int
// Number of days in 4 years
DateTime.DaysPer4Years = DateTime.DaysPerYear * 4 + 1; // 1461 //int
// Number of days in 100 years
DateTime.DaysPer100Years = DateTime.DaysPer4Years * 25 - 1; // 36524 //int
// Number of days in 400 years
DateTime.DaysPer400Years = DateTime.DaysPer100Years * 4 + 1; // 146097 //int
// Number of days from 1/1/0001 to 12/31/1600
DateTime.DaysTo1601 = DateTime.DaysPer400Years * 4; // 584388 //int
// Number of days from 1/1/0001 to 12/30/1899
DateTime.DaysTo1899 = DateTime.DaysPer400Years * 4 + DateTime.DaysPer100Years * 3 - 367; //int
// Number of days from 1/1/0001 to 12/31/1969
/**
 * @internal
 */
DateTime.DaysTo1970 = DateTime.DaysPer400Years * 4 + DateTime.DaysPer100Years * 3 + DateTime.DaysPer4Years * 17 + DateTime.DaysPerYear; // 719,162 //int
// Number of days from 1/1/0001 to 12/31/9999
DateTime.DaysTo10000 = DateTime.DaysPer400Years * 25 - 366; // 3652059 //int
/**
 *  @internal
 */
DateTime.MinTicks = 0; //long
/**
 *  @internal
 */
DateTime.MaxTicks = DateTime.DaysTo10000 * DateTime.TicksPerDay - 1; //long
DateTime.MaxMillis = DateTime.DaysTo10000 * DateTime.MillisPerDay; //long
DateTime.FileTimeOffset = DateTime.DaysTo1601 * DateTime.TicksPerDay; //long
DateTime.DoubleDateOffset = DateTime.DaysTo1899 * DateTime.TicksPerDay; //long
// The minimum OA date is 0100/01/01 (Note it's year 100).
// The maximum OA date is 9999/12/31
DateTime.OADateMinAsTicks = (DateTime.DaysPer100Years - DateTime.DaysPerYear) * DateTime.TicksPerDay; //long
// All OA dates must be greater than (not >=) OADateMinAsDouble
DateTime.OADateMinAsDouble = -657435.0; //double
// All OA dates must be less than (not <=) OADateMaxAsDouble
DateTime.OADateMaxAsDouble = 2958466.0; //double
DateTime.DatePartYear = 0; //int
DateTime.DatePartDayOfYear = 1; //int
DateTime.DatePartMonth = 2; //int
DateTime.DatePartDay = 3; //int
DateTime.DaysToMonth365 = [0, 31, 59, 90, 120, 151, 181, 212, 243, 273, 304, 334, 365];
DateTime.DaysToMonth366 = [0, 31, 60, 91, 121, 152, 182, 213, 244, 274, 305, 335, 366];
DateTime.MinValue = new DateTime(DateTime.MinTicks, DateTime_1.DateTimeKind.Unspecified);
DateTime.MaxValue = new DateTime(DateTime.MaxTicks, DateTime_1.DateTimeKind.Unspecified);
DateTime.TicksMask = 0x3FFFFFFFFFFFFFFF; //UInt64
DateTime.FlagsMask = 0xC000000000000000; //UInt64
DateTime.LocalMask = 0x8000000000000000; //UInt64
DateTime.TicksCeiling = 0x4000000000000000; //Int64
DateTime.KindUnspecified = 0x0000000000000000; //UInt64
DateTime.KindUtc = 0x4000000000000000; //UInt64
DateTime.KindLocal = 0x8000000000000000; //UInt64
DateTime.KindLocalAmbiguousDst = 0xC000000000000000; //UInt64
DateTime.KindShift = 62; //Int32
DateTime.TicksField = "ticks";
DateTime.DateDataField = "dateData";
exports.DateTime = DateTime;

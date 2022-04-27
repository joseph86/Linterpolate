# Linterpolate
2-dimensional interpolation add-in for Excel written in Visual Basic

This was originally written for the purpose of calculating the energy content of steam. It works with reasonable accuracy for that purpose and is perhaps a little better than straight nearest neighbor estimation. However, there is room for improvement.

Excel doesn't have a built in function dedicated to linear interpolation. It's still possible to do it, but it can be cumbersome. Normally, I would just complain and use the tools provided to muddle through, but I came upon a situation where writing a custom function was less work than creating the convoluted monster of standard excel functions required to do what I wanted to do, 2D linear interpolation.  Introducing LInterpolate.


Description

Returns the linear interpolation for the given new x. Finds the nearest neighbors in an array of known x's and returns the corresponding y interpolated from an array of known y's.

If the optional new y is also passed, returns the interpolated result from a 2D table of values defined by the intersection of the known x's columns and known y's rows.

Syntax

LInterpolate (x_range,  y_range, new_x,  [new_y ])

LInterpolate syntax has the following arguments:

x_range required. The range of known x's in your data used for interpolation.x_range must be a range of two or more values in ascending order for interpolation to work.
For predictable results, x_range should be a single contiguous range of values.
y_range required. The range of known y's in your data used for interpolation.y_range must be a range of two or more values.
if new_y is used, y_range must be in ascending order for interpolation to work.
For 2D interpolation, y_range must be perpendicular to x_range.  This means that if one is a single row, the other must be a single column.
For predictable results, y_range should be a single contiguous range of values.
new_x required. The new x used to interpolate the new y.new_x must be a single value or cell
new_y optional. The new y used for interpolating from a 2D table of data.if new_y is used, it must be a single value or cell

Remarks

I built LInterpolate to work with steam tables that have no gaps so LInterpolate will throw an error if it finds empty cells in the data. This is useful because it lets you know you've reached an edge case and alerts you to the possibility of spurious data. However I recognize that it's possible to have a 2D table that is not completely filled. If you find yourself in this situation, let me know, and I can help create a solution.

Examples

In the LInterpolateExamples.xlsx file, there are three sheets. One is the GNU General Public License explained below. The 'Example' sheet shows an example of single dimension linear interpolation. The code entered looks like this:

Linterpolate (F$7:BQ$7,  F$6:BQ$6,  A8)

The '2D Example' sheet is slightly more complicated. It interpolates values from a 2D table of the enthalpy of superheated steam at any pressure and temperature combination. The code entered looks like this:

Linterpolate (G$10:G$61,   H$9:BS$9,  A12,  B12)

Notice that the range of known x's is perpendicular to the range of known y's and the table of data used for interpolation is inferred from the spacial relationship of the two ranges. There are two cells where the LInterpolate function returned an error. This is desirable behavior for the reason I explained under Remarks.

Installation

One-time use

For one-time use of LInterpolate, just open the LInterpolate.xla file. As long as the file is open, you can use LInterpolate in other spreadsheets. Use the 'Insert Function' dialog box to avoid typos.

Available to all spreadsheets

If you want the function available every time you open excel you'll need to copy LInterpolate.xla to your XLSTART folder. The path should be something like

C:\Users\[username]\Appdata\Roaming\Microsoft\Excel\XLSTART

Letting others use LInterpolate

If you use LInterpolate and then send your spreadsheet so someone else, they'll only see the results you calculated with it, but they will not be able to recalculate the results. If you want others to have access to LInterpolate, you'll have to either:

1) send them the LInterpolate.xla in addition to your spreadsheet, or
2) embed the function in the spreadsheet

Embedding in the spreadsheet

The easiest way to make sure LInterpolate will work no matter where your spreadsheet goes is to use a copy of LInterpolate.xla as a template to make your spreadsheet. Save it as a new name with the .xlsm extension to enable macros.

If you already have a spreadsheet you would like to augment with LInterpolate, follow one of the many online tutorials for creating user-defined functions. Open the code editor with Alt-F11. Just copy and paste the LInterpolate code and save.

License

This program is distributed under the terms of the GNU General Public License v3

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program. If not, see GNU Licenses.

Final Notes

This saved my company lots of time and money in processing excel data. If you find it useful, please Donate Bitcoins. Just copy my BTC address

1BQmUrozShAccPnKvihf6486UiGStxqJ7G

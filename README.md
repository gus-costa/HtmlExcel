HtmlExcel
==========


Turn HTML tables into multi-sheet Excel files.


Example
-------

```php
require_once('HtmlExcel.php');

$css = "
.red {
	color: red;
}";

$numbers = '<table>
<tr>
	<td class="red">1</td>
	<td>2</td>
	<td>3</td>
</tr>
<tr>
	<td>4</td>
	<td class="red">5</td>
	<td>6</td>
</tr>
<tr>
	<td>7</td>
	<td>8</td>
	<td class="red">9</td>
</tr>
</table>';

$names = '<table>
  <tr>
    <th>First name</th>
    <th>Last name</th>
  </tr>
  <tr>
    <td>John</td>
    <td>Doe</td>
  </tr>
  <tr>
    <td>Jane</td>
    <td>Doe</td>
  </tr>
</table>';

$xls = new HtmlExcel();
$xls->setCss($css);
$xls->addSheet("Numbers", $numbers);
$xls->addSheet("Names", $names);
$xls->headers();
echo $xls->buildFile();
```
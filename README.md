# NpoiExcelHelper
npoi相关的东西写了几个月，发现类似的代码还是没有的；
为了后面的人能少走一些弯路，所有分享了以下代码。
英语不是太好，就不做三脚猫翻译，下面列举目前实现功能：
注意：源码中的备注都是中文

1，setAllSheetToAuto：
假设有一个单元格的公式为=A1+B1，那么修改A1或B1的值后，将该workbook保存为文件，会发现公式中的结果依旧是之前的内容，并没有根据A1与B1计算新的单元格结果；
此时就需要使用这个函数，调用后可以使workbook得公式自动运算：其中核心为：workbook.GetSheetAt(i).ForceFormulaRecalculation = true;

2，getAllMergedRegions：
遍历一个sheet中的所有合并单元格；

3，DelRows：
删除一行或多行：
众所周知，npoi是没有删除的函数，唯一能够起到类似作用的就是shiftrow，但实际当被删除的行中包含合并单元格时会造成该sheet错乱；
这个函数就是为了克服这个问题而编写；
其中核心思想：首先遍历需要删除的行，找出其中全部合并单元格，并将这些单元格逐个删除。
最后进行shiftrow上移

4，CopyRow：
同一个sheet内复制一行到目标位置；
比较早期编写的函数，其中参考到网上的资料，名字忘了不好意思。

5，CopyRows：
同一个sheet内复制多行到目标位置；
copyRow的升级版；其中算法缺陷较大，当在本sheet中有越多的合并单元格时耗费的时间将会越多。
原理是调用CopyRowWithoutMergedRegion进行复制，复制完毕后再将源rows中的合并单元格样式进行复制。

6，CopyRowWithoutMergedRegion：
复制一行而不复制合并单元格的样式；

7，CopyRowOverSheet：
后期编写函数，实现从sheet A复制一行到sheet B目标行

8，CopyRowsOverSheet：
最常用函数，跨sheet复制多行到目标sheet的某一行

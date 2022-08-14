# FakeXlsx
Porting of Gist to C#
https://gist.github.com/iso2022jp/721df3095f4df512bfe2327503ea1119

### Additional Functions
- Create a title line from a property name.
- Fixed display of the first line.
- Add simplified AutoFitColumns function.

### Benchmark

BenchmarkDotNet=v0.13.1, OS=Windows 10.0.22000
Intel Core i7-10610U CPU 1.80GHz, 1 CPU, 8 logical and 4 physical cores
.NET SDK=6.0.400
  [Host]     : .NET 6.0.8 (6.0.822.36306), X64 RyuJIT  [AttachedDebugger]
  DefaultJob : .NET 6.0.8 (6.0.822.36306), X64 RyuJIT

|                 Method |      N |       Mean |      Error |     StdDev | Ratio | RatioSD |      Gen 0 |     Gen 1 |  Allocated |
|----------------------- |------- |-----------:|-----------:|-----------:|------:|--------:|-----------:|----------:|-----------:|
|        ReflectionAsync |   1000 |   4.486 ms |  0.0844 ms |  0.0790 ms |  1.00 |    0.00 |   320.3125 |   15.6250 |   1,323 KB |
|    ExpressionTreeAsync |   1000 |   4.689 ms |  0.0919 ms |  0.1227 ms |  1.04 |    0.02 |   320.3125 |   62.5000 |   1,323 KB |
|  ExpressionTreeOpAsync |   1000 |   4.605 ms |  0.0676 ms |  0.0632 ms |  1.03 |    0.02 |   187.5000 |    7.8125 |     772 KB |
| ExpressionTreeOp2Async |   1000 |   3.607 ms |  0.0483 ms |  0.0452 ms |  0.80 |    0.02 |   187.5000 |    3.9063 |     770 KB |
|                        |        |            |            |            |       |         |            |           |            |
|        ReflectionAsync | 100000 | 665.889 ms | 10.0648 ms |  8.9222 ms |  1.00 |    0.00 | 21000.0000 | 5000.0000 | 132,072 KB |
|    ExpressionTreeAsync | 100000 | 654.778 ms | 13.0358 ms | 12.1937 ms |  0.98 |    0.02 | 21000.0000 | 5000.0000 | 132,072 KB |
|  ExpressionTreeOpAsync | 100000 | 620.747 ms |  9.7818 ms |  8.6713 ms |  0.93 |    0.02 | 12000.0000 | 3000.0000 |  74,280 KB |
| ExpressionTreeOp2Async | 100000 | 541.910 ms | 10.6133 ms | 10.8991 ms |  0.81 |    0.02 | 12000.0000 | 3000.0000 |  74,250 KB |


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
.NET SDK=6.0.301
  [Host]     : .NET 6.0.6 (6.0.622.26707), X64 RyuJIT  [AttachedDebugger]
  DefaultJob : .NET 6.0.6 (6.0.622.26707), X64 RyuJIT


|                 Method |      N |       Mean |      Error |     StdDev |     Median | Ratio | RatioSD |      Gen 0 |     Gen 1 |  Allocated |
|----------------------- |------- |-----------:|-----------:|-----------:|-----------:|------:|--------:|-----------:|----------:|-----------:|
|        ReflectionAsync |   1000 |   6.639 ms |  0.2987 ms |  0.8713 ms |   6.305 ms |  1.00 |    0.00 |   320.3125 |   23.4375 |   1,328 KB |
|    ExpressionTreeAsync |   1000 |   5.764 ms |  0.0762 ms |  0.0713 ms |   5.733 ms |  0.88 |    0.06 |   320.3125 |   23.4375 |   1,328 KB |
|  ExpressionTreeOpAsync |   1000 |   5.591 ms |  0.0629 ms |  0.0526 ms |   5.587 ms |  0.85 |    0.06 |   187.5000 |         - |     777 KB |
| ExpressionTreeOp2Async |   1000 |   4.822 ms |  0.0960 ms |  0.0943 ms |   4.792 ms |  0.73 |    0.05 |   187.5000 |   39.0625 |     776 KB |
|                        |        |            |            |            |            |       |         |            |           |            |
|        ReflectionAsync | 100000 | 636.028 ms |  9.6319 ms |  8.5384 ms | 633.882 ms |  1.00 |    0.00 | 21000.0000 | 5000.0000 | 132,078 KB |
|    ExpressionTreeAsync | 100000 | 630.228 ms | 11.8150 ms | 12.6419 ms | 631.758 ms |  0.99 |    0.02 | 21000.0000 | 5000.0000 | 132,076 KB |
|  ExpressionTreeOpAsync | 100000 | 601.579 ms |  9.6817 ms |  8.5825 ms | 599.061 ms |  0.95 |    0.02 | 12000.0000 | 3000.0000 |  74,280 KB |
| ExpressionTreeOp2Async | 100000 | 531.662 ms |  9.9531 ms | 11.0629 ms | 531.445 ms |  0.84 |    0.02 | 12000.0000 | 3000.0000 |  74,256 KB |


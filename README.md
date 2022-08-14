# FakeXlsx
Porting of Gist to C#
https://gist.github.com/iso2022jp/721df3095f4df512bfe2327503ea1119

### Additional Functions
- Create a title line from a property name.
- Fixed display of the first line.
- Add simplified AutoFitColumns function.

### Benchmark

// * Summary *

BenchmarkDotNet=v0.13.1, OS=Windows 10.0.22000
Intel Core i7-10610U CPU 1.80GHz, 1 CPU, 8 logical and 4 physical cores
.NET SDK=6.0.400
  [Host]     : .NET 6.0.8 (6.0.822.36306), X64 RyuJIT  [AttachedDebugger]
  DefaultJob : .NET 6.0.8 (6.0.822.36306), X64 RyuJIT

|                 Method |      N |       Mean |      Error |     StdDev | Ratio | RatioSD |      Gen 0 |     Gen 1 |  Allocated |
|----------------------- |------- |-----------:|-----------:|-----------:|------:|--------:|-----------:|----------:|-----------:|
|        ReflectionAsync |   1000 |   6.402 ms |  0.0815 ms |  0.0681 ms |  1.00 |    0.00 |   351.5625 |   70.3125 |   1,450 KB |
|    ExpressionTreeAsync |   1000 |   6.341 ms |  0.1239 ms |  0.1377 ms |  1.00 |    0.03 |   351.5625 |   70.3125 |   1,450 KB |
|  ExpressionTreeOpAsync |   1000 |   6.077 ms |  0.1158 ms |  0.1137 ms |  0.95 |    0.02 |   195.3125 |    7.8125 |     836 KB |
| ExpressionTreeOp2Async |   1000 |   4.673 ms |  0.0907 ms |  0.1114 ms |  0.74 |    0.02 |   203.1250 |    7.8125 |     833 KB |
|                        |        |            |            |            |       |         |            |           |            |
|        ReflectionAsync | 100000 | 799.193 ms | 15.6398 ms | 16.0610 ms |  1.00 |    0.00 | 23000.0000 | 6000.0000 | 144,000 KB |
|    ExpressionTreeAsync | 100000 | 773.797 ms | 10.0454 ms |  8.9050 ms |  0.97 |    0.03 | 23000.0000 | 6000.0000 | 144,001 KB |
|  ExpressionTreeOpAsync | 100000 | 731.430 ms | 13.0427 ms | 11.5620 ms |  0.92 |    0.02 | 13000.0000 | 3000.0000 |  80,079 KB |
| ExpressionTreeOp2Async | 100000 | 602.000 ms | 10.5444 ms |  9.3473 ms |  0.75 |    0.02 | 13000.0000 | 3000.0000 |  79,925 KB |


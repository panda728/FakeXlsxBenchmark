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


|                 Method |      N |       Mean |      Error |    StdDev | Ratio | RatioSD |      Gen 0 |     Gen 1 |  Allocated |
|----------------------- |------- |-----------:|-----------:|----------:|------:|--------:|-----------:|----------:|-----------:|
|        ReflectionAsync |   1000 |   4.621 ms |  0.0791 ms | 0.0740 ms |  1.00 |    0.00 |   320.3125 |   15.6250 |   1,323 KB |
|    ExpressionTreeAsync |   1000 |   5.157 ms |  0.1013 ms | 0.1206 ms |  1.11 |    0.03 |   320.3125 |   62.5000 |   1,323 KB |
|  ExpressionTreeOpAsync |   1000 |   4.892 ms |  0.0579 ms | 0.0541 ms |  1.06 |    0.02 |   187.5000 |    7.8125 |     772 KB |
| ExpressionTreeOp2Async |   1000 |   3.761 ms |  0.0432 ms | 0.0383 ms |  0.82 |    0.02 |   187.5000 |    3.9063 |     770 KB |
|                        |        |            |            |           |       |         |            |           |            |
|        ReflectionAsync | 100000 | 659.827 ms | 10.6794 ms | 8.9178 ms |  1.00 |    0.00 | 21000.0000 | 5000.0000 | 132,073 KB |
|    ExpressionTreeAsync | 100000 | 655.398 ms |  9.0048 ms | 7.5194 ms |  0.99 |    0.02 | 21000.0000 | 5000.0000 | 132,072 KB |
|  ExpressionTreeOpAsync | 100000 | 620.269 ms |  8.8600 ms | 8.2876 ms |  0.94 |    0.02 | 12000.0000 | 3000.0000 |  74,278 KB |
| ExpressionTreeOp2Async | 100000 | 524.755 ms |  7.8092 ms | 6.9226 ms |  0.80 |    0.02 | 12000.0000 | 3000.0000 |  74,252 KB |


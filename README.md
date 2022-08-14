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
  [Host]   : .NET 6.0.8 (6.0.822.36306), X64 RyuJIT  [AttachedDebugger]
  ShortRun : .NET 6.0.8 (6.0.822.36306), X64 RyuJIT

Job=ShortRun  IterationCount=3  LaunchCount=1
WarmupCount=3

|                 Method |      N |       Mean |       Error |     StdDev | Ratio | RatioSD |      Gen 0 |     Gen 1 |  Allocated |
|----------------------- |------- |-----------:|------------:|-----------:|------:|--------:|-----------:|----------:|-----------:|
|        ReflectionAsync |   1000 |   4.382 ms |   1.5948 ms |  0.0874 ms |  1.00 |    0.00 |   320.3125 |   15.6250 |   1,323 KB |
|    ExpressionTreeAsync |   1000 |   4.449 ms |   0.5913 ms |  0.0324 ms |  1.02 |    0.01 |   320.3125 |   15.6250 |   1,323 KB |
|  ExpressionTreeOpAsync |   1000 |   4.304 ms |   0.7163 ms |  0.0393 ms |  0.98 |    0.02 |   187.5000 |    7.8125 |     773 KB |
| ExpressionTreeOp2Async |   1000 |   3.560 ms |   0.4516 ms |  0.0248 ms |  0.81 |    0.01 |   187.5000 |    3.9063 |     770 KB |
|                        |        |            |             |            |       |         |            |           |            |
|        ReflectionAsync | 100000 | 659.245 ms |  90.3217 ms |  4.9508 ms |  1.00 |    0.00 | 21000.0000 | 5000.0000 | 132,074 KB |
|    ExpressionTreeAsync | 100000 | 652.458 ms | 216.0968 ms | 11.8450 ms |  0.99 |    0.02 | 21000.0000 | 5000.0000 | 132,074 KB |
|  ExpressionTreeOpAsync | 100000 | 609.833 ms |  85.2152 ms |  4.6709 ms |  0.93 |    0.01 | 12000.0000 | 3000.0000 |  74,406 KB |
| ExpressionTreeOp2Async | 100000 | 522.271 ms |  34.8645 ms |  1.9110 ms |  0.79 |    0.01 | 12000.0000 | 3000.0000 |  74,251 KB |


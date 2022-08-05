# FakeXlsx
Porting of Gist to C#
https://gist.github.com/iso2022jp/721df3095f4df512bfe2327503ea1119

### Additional Functions
- Create a title line from a property name.
- Fixed display of the first line.
- Add simplified AutoFit function.

### Benchmark

BenchmarkDotNet=v0.13.1, OS=Windows 10.0.22000
Intel Core i7-10610U CPU 1.80GHz, 1 CPU, 8 logical and 4 physical cores
.NET SDK=6.0.301
  [Host]     : .NET 6.0.6 (6.0.622.26707), X64 RyuJIT  [AttachedDebugger]
  DefaultJob : .NET 6.0.6 (6.0.622.26707), X64 RyuJIT


|                Method |      N |       Mean |      Error |     StdDev | Ratio | RatioSD |      Gen 0 |     Gen 1 |  Allocated |
|---------------------- |------- |-----------:|-----------:|-----------:|------:|--------:|-----------:|----------:|-----------:|
|       ReflectionAsync |   1000 |   5.216 ms |  0.0772 ms |  0.0685 ms |  1.00 |    0.00 |   328.1250 |    7.8125 |   1,344 KB |
|   ExpressionTreeAsync |   1000 |   6.887 ms |  0.0789 ms |  0.0738 ms |  1.32 |    0.03 |   335.9375 |   93.7500 |   1,395 KB |
| ExpressionTreeOpAsync |   1000 |   6.465 ms |  0.0803 ms |  0.0751 ms |  1.24 |    0.02 |   187.5000 |   15.6250 |     783 KB |
|                       |        |            |            |            |       |         |            |           |            |
|       ReflectionAsync | 100000 | 614.567 ms | 11.3629 ms | 10.6289 ms |  1.00 |    0.00 | 21000.0000 | 5000.0000 | 132,096 KB |
|   ExpressionTreeAsync | 100000 | 626.500 ms | 12.4696 ms | 30.5882 ms |  1.03 |    0.05 | 21000.0000 | 5000.0000 | 132,143 KB |
| ExpressionTreeOpAsync | 100000 | 549.215 ms |  9.6122 ms |  8.9913 ms |  0.89 |    0.02 | 11000.0000 | 3000.0000 |  68,872 KB |

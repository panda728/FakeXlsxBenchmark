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

|                 Method |      N |        Mean |     Error |    StdDev | Ratio | RatioSD |      Gen 0 |     Gen 1 |  Allocated |
|----------------------- |------- |------------:|----------:|----------:|------:|--------:|-----------:|----------:|-----------:|
|        ReflectionAsync |   1000 |    12.67 ms |  0.462 ms |  1.341 ms |  1.00 |    0.00 |   312.5000 |   46.8750 |   1,328 KB |
|    ExpressionTreeAsync |   1000 |    13.34 ms |  0.628 ms |  1.802 ms |  1.06 |    0.19 |   312.5000 |   31.2500 |   1,328 KB |
|  ExpressionTreeOpAsync |   1000 |    12.25 ms |  0.448 ms |  1.265 ms |  0.98 |    0.16 |   187.5000 |         - |     777 KB |
| ExpressionTreeOp2Async |   1000 |    10.84 ms |  0.436 ms |  1.237 ms |  0.86 |    0.12 |   187.5000 |   46.8750 |     776 KB |
|                        |        |             |           |           |       |         |            |           |            |
|        ReflectionAsync | 100000 | 1,087.91 ms | 21.056 ms | 32.155 ms |  1.00 |    0.00 | 21000.0000 | 5000.0000 | 132,076 KB |
|    ExpressionTreeAsync | 100000 | 1,076.49 ms | 21.459 ms | 38.144 ms |  0.99 |    0.05 | 21000.0000 | 5000.0000 | 132,076 KB |
|  ExpressionTreeOpAsync | 100000 | 1,026.86 ms | 20.527 ms | 43.298 ms |  0.94 |    0.05 | 12000.0000 | 5000.0000 |  74,280 KB |
| ExpressionTreeOp2Async | 100000 |   936.37 ms | 18.273 ms | 31.030 ms |  0.86 |    0.04 | 12000.0000 | 3000.0000 |  74,256 KB |

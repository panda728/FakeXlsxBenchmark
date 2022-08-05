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

|                Method |      N |       Mean |      Error |    StdDev | Ratio | RatioSD |      Gen 0 |     Gen 1 |  Allocated |
|---------------------- |------- |-----------:|-----------:|----------:|------:|--------:|-----------:|----------:|-----------:|
|       ReflectionAsync |   1000 |   5.163 ms |  0.1004 ms | 0.0986 ms |  1.00 |    0.00 |   320.3125 |   23.4375 |   1,328 KB |
|   ExpressionTreeAsync |   1000 |   5.373 ms |  0.1072 ms | 0.1700 ms |  1.03 |    0.05 |   320.3125 |   23.4375 |   1,328 KB |
| ExpressionTreeOpAsync |   1000 |   5.389 ms |  0.0769 ms | 0.0682 ms |  1.04 |    0.02 |   187.5000 |         - |     777 KB |
|                       |        |            |            |           |       |         |            |           |            |
|       ReflectionAsync | 100000 | 609.558 ms | 10.0962 ms | 8.9500 ms |  1.00 |    0.00 | 21000.0000 | 5000.0000 | 132,076 KB |
|   ExpressionTreeAsync | 100000 | 594.267 ms | 10.0561 ms | 7.8512 ms |  0.98 |    0.02 | 21000.0000 | 5000.0000 | 132,076 KB |
| ExpressionTreeOpAsync | 100000 | 578.450 ms |  9.6359 ms | 9.0134 ms |  0.95 |    0.02 | 12000.0000 | 3000.0000 |  74,280 KB |

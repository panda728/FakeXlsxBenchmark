using BenchmarkDotNet.Running;
using FakeXlsxBenchmark;

#if DEBUG
var test = new BuilderTest();
test.Setup();
await test.NormalAsync();
#else
var summary = BenchmarkRunner.Run<BuilderTest>();
#endif

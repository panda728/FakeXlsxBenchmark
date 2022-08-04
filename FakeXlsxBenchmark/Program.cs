using BenchmarkDotNet.Running;
using FakeXlsxBenchmark;

#if DEBUG
var test = new BuilderTest();
test.Setup();
//await test.ReflectionAsync();
//await test.ExpressionTreeAsync();
await test.ExpressionTreeOpAsync();

#else
var summary = BenchmarkRunner.Run<BuilderTest>();
#endif

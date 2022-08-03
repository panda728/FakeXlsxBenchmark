using BenchmarkDotNet.Running;
using FakeXlsxBenchmark;

#if DEBUG
var test = new BuilderTest();
test.Setup();
//await test.ReflectionAsync();
test.ExpressionTree();

#else
var summary = BenchmarkRunner.Run<BuilderTest>();
#endif

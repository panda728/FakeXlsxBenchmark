using BenchmarkDotNet.Running;
using FakeXlsxBenchmark;

#if DEBUG


var test = new BuilderTest();
test.Setup();

await test.ReflectionAsync();
await test.ExpressionTreeAsync();
await test.ExpressionTreeOpAsync();
await test.ExpressionTreeOp2Async();

var builder = new FakeExcel.Builder();
builder.CreateExcelFile(@"test\\test.xlsx", new string[] { "test1", "test2" });

var titles = new string[] { "Id", "名", "姓", "氏名", "ユーザー名", "Email", "ユニークキー", "Guid", "プロフィール画像", "カートGuid", "TEL", "作成日時", "性別", "注文" };
builder.CreateExcelFile(@"test\\builder.xlsx", test.Users, showTitleRow: true, columnAutoFit: true, titles: titles);

#else
var summary = BenchmarkRunner.Run<BuilderTest>();
#endif

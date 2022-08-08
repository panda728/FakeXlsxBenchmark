using BenchmarkDotNet.Attributes;
using Bogus;
using FakeExcelBuilder.ExpressionTree;
using FakeExcelBuilder.Reflection;
using static Bogus.DataSets.Name;

namespace FakeXlsxBenchmark
{
    [MarkdownExporterAttribute.GitHub]
    [ShortRunJob]
    [MemoryDiagnoser]
    public class BuilderTest
    {
        readonly FakeExcelBuilder.Reflection.Builder _builderRef = new();
        readonly FakeExcelBuilder.ExpressionTree.Builder _builderExp = new();
        readonly FakeExcelBuilder.ExpressionTreeOp.Builder _builderExpOp = new();
        readonly FakeExcelBuilder.ExpressionTreeOp.Builder _builderExpOp2 = new();
        List<User>? _users;
        public BuilderTest()
        {
        }

        [Params(1000, 100000)]
        public int N = 1000;

        [GlobalSetup]
        public void Setup()
        {
            //_builderRef.Compile(typeof(User));
            //_builderExp.Compile(typeof(User));
            //_builderExpOp.Compile(typeof(User));

            Randomizer.Seed = new Random(8675309);

            var fruit = new[] { "apple", "banana", "orange", "strawberry", "kiwi" };

            var orderIds = 0;
            var testOrders = new Faker<Order>()
                .StrictMode(true)
                .RuleFor(o => o.OrderId, f => orderIds++)
                .RuleFor(o => o.Item, f => f.PickRandom(fruit))
                .RuleFor(o => o.Quantity, f => f.Random.Number(1, 10))
                .RuleFor(o => o.LotNumber, f => f.Random.Int(0, 100).OrNull(f, .8f));

            var userIds = 0;
            var testUsers = new Faker<User>()
                .CustomInstantiator(f => new User(userIds++, f.Random.Replace("###-##-####")))
                .RuleFor(u => u.Gender, f => f.PickRandom<Gender>())
                .RuleFor(u => u.FirstName, (f, u) => f.Name.FirstName(u.Gender))
                .RuleFor(u => u.LastName, (f, u) => f.Name.LastName(u.Gender))
                .RuleFor(u => u.Avatar, f => f.Internet.Avatar())
                .RuleFor(u => u.UserName, (f, u) => f.Internet.UserName(u.FirstName, u.LastName))
                .RuleFor(u => u.Email, (f, u) => f.Internet.Email(u.FirstName, u.LastName))
                .RuleFor(u => u.SomethingUnique, f => $"Value {f.UniqueIndex}")
                .RuleFor(u => u.CreateTime, f => DateTime.Now)
                .RuleFor(u => u.SomeGuid, f => Guid.NewGuid())
                .RuleFor(u => u.CartId, f => Guid.NewGuid())
                .RuleFor(u => u.FullName, (f, u) => u.FirstName + " " + u.LastName)
                .RuleFor(u => u.Orders, f => testOrders.Generate(3).ToList())
                .FinishWith((f, u) =>
                {
                    //Console.WriteLine("User Created! Id={0}", u.Id);
                });

            _users = testUsers.Generate(N);
            if(!Directory.Exists("test"))
                Directory.CreateDirectory("test");
        }

        [Benchmark(Baseline = true)]
        public async Task ReflectionAsync()
        {
            if (_users == null)
                throw new ApplicationException("users is null");
            var fileName = @"test\Reflection.xlsx";
            await _builderRef.RunAsync(fileName, _users);
        }

        [Benchmark]
        public async Task ExpressionTreeAsync()
        {
            if (_users == null)
                throw new ApplicationException("users is null");
            var fileName = @"test\ExpressionTree.xlsx";
            await _builderExp.RunAsync(fileName, _users);
        }
        [Benchmark]
        public async Task ExpressionTreeOpAsync()
        {
            if (_users == null)
                throw new ApplicationException("users is null");
            var fileName = @"test\ExpressionTreeOp.xlsx";
            await _builderExpOp.RunAsync(fileName, _users);
        }
        [Benchmark]
        public async Task ExpressionTreeOp2Async()
        {
            if (_users == null)
                throw new ApplicationException("users is null");
            var fileName = @"test\ExpressionTreeOp2.xlsx";
            await _builderExpOp2.RunAsync(fileName, _users);
        }
    }
}

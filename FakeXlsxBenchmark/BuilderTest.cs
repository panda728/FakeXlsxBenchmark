using BenchmarkDotNet.Attributes;
using Bogus;
using static Bogus.DataSets.Name;

namespace FakeXlsxBenchmark
{
    [MarkdownExporterAttribute.GitHub]
    [ShortRunJob]
    [MemoryDiagnoser]
    public class BuilderTest
    {
        readonly FakeExcelBuilder.Builder _builder;
        List<User>? _users;
        public BuilderTest()
        {
            _builder = new FakeExcelBuilder.Builder();
        }

        [Params(1000, 100000)]
        public int N = 1000;

        [GlobalSetup]
        public void Setup()
        {
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
        }

        [Benchmark(Baseline = true)]
        public async Task NormalAsync()
        {
            if (_users == null)
                throw new ApplicationException("users is null");
            var fileName = @"test\hellowworld.xlsx";
            await _builder.RunAsync(fileName, _users);
        }
    }
}

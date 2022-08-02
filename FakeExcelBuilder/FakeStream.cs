namespace FakeExcelBuilder
{
    public class FakeStream : MemoryStream
    {
        public override void Write(byte[] buffer, int offset, int count)
        {
            //base.Write(buffer, offset, count);
        }

        public override void WriteByte(byte value)
        {
            //base.WriteByte(value);
        }

        public override Task WriteAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
        {
            //return base.WriteAsync(buffer, offset, count, cancellationToken);
            return Task.CompletedTask;
        }

        public override ValueTask WriteAsync(ReadOnlyMemory<byte> buffer, CancellationToken cancellationToken = default)
        {
            //return base.WriteAsync(buffer, cancellationToken);
            return ValueTask.CompletedTask;
        }

        public override void Write(ReadOnlySpan<byte> buffer)
        {
            //base.Write(buffer);
        }
    }
}

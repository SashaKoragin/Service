using System;
using System.Globalization;

namespace OpenPop.Pop3
{
	public abstract class Disposable : IDisposable
	{
		protected bool IsDisposed
		{
			get;
			private set;
		}

		~Disposable()
		{
			Dispose(disposing: false);
		}

		public void Dispose()
		{
			if (!IsDisposed)
			{
				try
				{
					Dispose(disposing: true);
				}
				finally
				{
					IsDisposed = true;
					GC.SuppressFinalize(this);
				}
			}
		}

		protected virtual void Dispose(bool disposing)
		{
		}

		protected void AssertDisposed()
		{
			if (IsDisposed)
			{
				string fullName = GetType().FullName;
				throw new ObjectDisposedException(fullName, string.Format(CultureInfo.InvariantCulture, "Cannot access a disposed {0}.", new object[1]
				{
					fullName
				}));
			}
		}
	}
}

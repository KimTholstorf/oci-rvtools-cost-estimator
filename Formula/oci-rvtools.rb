class OciRvtools < Formula
  desc "Convert RVTools Excel exports into an Oracle Cloud (OCI) monthly cost estimate workbook"
  homepage "https://github.com/KimTholstorf/oci-rvtools-cost-estimator"
  url "https://files.pythonhosted.org/packages/13/60/47dfffe63c9f6a810e32993ab4fca934976098e3c6b74e3c604a30ad7d4f/oci_rvtools-1.2.0.tar.gz"
  sha256 "6ddc3c7e87de1c8e60b488d5289af34a86d301e2c9fc256dd63304860167d5da"
  license "MIT"

  depends_on "python3"

  def install
    system "python3", "-m", "venv", libexec
    system libexec/"bin/pip", "install", "--no-cache-dir", "oci-rvtools==#{version}"
    bin.install_symlink libexec/"bin/oci-rvtools"
  end

  test do
    system bin/"oci-rvtools", "--version"
  end
end

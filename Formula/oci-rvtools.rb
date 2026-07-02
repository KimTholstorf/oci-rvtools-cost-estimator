class OciRvtools < Formula
  desc "Convert RVTools Excel exports into an Oracle Cloud (OCI) monthly cost estimate workbook"
  homepage "https://github.com/KimTholstorf/oci-rvtools-cost-estimator"
  url "https://files.pythonhosted.org/packages/54/e9/de93330772356260778e2196b02dfd44efb62b3ef3c37f18e095e5ddc20c/oci_rvtools-1.3.0.tar.gz"
  sha256 "520502249c4002f734b1c371a2ad30fa32361c22c9cb5a3c9f829eafc1e2cefb"
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

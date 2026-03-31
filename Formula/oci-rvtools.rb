class OciRvtools < Formula
  desc "Convert RVTools Excel exports into an Oracle Cloud (OCI) monthly cost estimate workbook"
  homepage "https://github.com/KimTholstorf/oci-rvtools-cost-estimator"
  url "https://files.pythonhosted.org/packages/source/o/oci-rvtools/oci_rvtools-1.0.7.tar.gz"
  sha256 ""
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

class OciRvtools < Formula
  desc "Convert RVTools Excel exports into an Oracle Cloud (OCI) monthly cost estimate workbook"
  homepage "https://github.com/KimTholstorf/oci-rvtools-cost-estimator"
  url "https://files.pythonhosted.org/packages/44/ee/0b8eae3d3fb011aa378446dac4d8deffaee7ed452a0b7094d70b1607b092/oci_rvtools-1.1.0.tar.gz"
  sha256 "c399b392205e7573ad048b128d3d581e287ac7fc7385a749c5a109cec02020e5"
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

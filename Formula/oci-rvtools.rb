class OciRvtools < Formula
  desc "Convert RVTools Excel exports into an Oracle Cloud (OCI) monthly cost estimate workbook"
  homepage "https://github.com/KimTholstorf/oci-rvtools-cost-estimator"
  url "https://files.pythonhosted.org/packages/d2/11/db56aa8f2452a93e431a7337d80c43203429605f05816efdac016b5ba001/oci_rvtools-1.0.10.tar.gz"
  sha256 "d32b382887b58b124f9151dd6aa5da23dcb9494d776c9140c27b7cf5f9e90b19"
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

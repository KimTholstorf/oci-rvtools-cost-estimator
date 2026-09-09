class OciRvtools < Formula
  desc "Convert RVTools Excel exports into an Oracle Cloud (OCI) monthly cost estimate workbook"
  homepage "https://github.com/KimTholstorf/oci-rvtools-cost-estimator"
  url "https://files.pythonhosted.org/packages/e3/6a/99b2b2e419070d14cca1d444e176e835531a6afb1f7c11f3a801cf00854b/oci_rvtools-1.3.1.tar.gz"
  sha256 "8fda2f0df5f5d6e2a117c5700b4940f56cd9300397a3d008f94d4c133478c695"
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

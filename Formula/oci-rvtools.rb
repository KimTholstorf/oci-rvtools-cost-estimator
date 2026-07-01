class OciRvtools < Formula
  desc "Convert RVTools Excel exports into an Oracle Cloud (OCI) monthly cost estimate workbook"
  homepage "https://github.com/KimTholstorf/oci-rvtools-cost-estimator"
  url "https://files.pythonhosted.org/packages/f6/59/a296d41a0a655407117e630f696be2c15efbb837371ec739b9cc9aecc74e/oci_rvtools-1.2.2.tar.gz"
  sha256 "fe7da5993afbad93881e34b9d3ea4234b24b4e60c4e659d357f6eba532fe089b"
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

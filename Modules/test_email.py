import pytest
from draftEmail import ValidateEmail

# check email inputs
def test_email1():
    email = "myname326@gmail.com"
    assert ValidateEmail.check_email(email) is True

def test_email2():
  email = "my.ownsite@acn.org"
  assert ValidateEmail.check_email(email) is True

def test_email3():
  email = "myname326.com"
  assert ValidateEmail.check_email(email) is False

def test_email4():
  email = "myname326@.com"
  assert ValidateEmail.check_email(email) is False

if __name__ == "__main__":
    pytest.main()
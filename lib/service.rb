require 'rubygems'
require 'win32/daemon'
require 'selenium-webdriver'
require 'mail'
require 'json'
include Win32

class WebUiDaemon < Daemon

  def service_main
    log 'started'
    while running?
      log 'running'

      configure_smtp

      has_run = false
      loop do
        if Time.now.hour == 0 and not has_run
          register_user 'Joey'
          register_user 'Katie'
          register_user 'Sabra'
          register_user 'Judy'
          register_user 'Gary'
          has_run = true
        elsif Time.now.hour == 23 and has_run
          # Reset has_run so that automation is picked up the following day.
          has_run = false
        end
        sleep 60
      end
    end
  end

  def service_stop
    log 'ended'
    exit!
  end

  def log(text)
    File.open('log.txt', 'a') { |f| f.puts "#{Time.now}: #{text}" }
  end

  private

  def configure_smtp
    begin
      config = JSON.load(File.open('hgtv_config.txt'))
      password = config['gmail_password']

      options = { :address              => "smtp.gmail.com",
                  :port                 => 587,
                  :domain               => 'GhostGear.com',
                  :user_name            => 'ghostgearmk.i@gmail.com',
                  :password             => password,
                  :authentication       => 'plain',
                  :enable_starttls_auto => true  }

      Mail.defaults do
        delivery_method :smtp, options
      end
    rescue => ex
      log ex
    end
  end

  def register_user(user)
    begin
      hgtv_dream_register user
      log "Successfully registered #{user} for HGTV dream home."
    rescue => ex
      log "WebUi service encountered the following error when registering #{user}. Ex:\n#{ex.message}"
      ex.backtrace.each { |line| log line }

      Mail.deliver do
        to 'kangaroo128@yahoo.com'
        from 'ghostgearmk.i@gmail.com'
        subject 'WebUi HGTV Error'
        body "WebUi service encountered the following error when registering #{user}.\nEx:\n#{ex.message}\nBacktrace:\n#{ex.backtrace}"
      end
    end  
  end

  def hgtv_dream_register(user)
    driver = nil
    ex = nil

    config = JSON.load(File.open('hgtv_config.txt'))
    user_config = config['users'].select{ |u| u['first_name'] == user}.first

    email = user_config['email']
    first_name = user_config['first_name']
    last_name = user_config['last_name']
    address = user_config['address']
    city = user_config['city']
    state = user_config['state']
    zip = user_config['zip']
    phone = user_config['phone']
    gender = user_config['gender']
    birth_month = user_config['birth_month']
    birth_day = user_config['birth_day']
    birth_year = user_config['birth_year']
    cable_provider = user_config['cable_provider']

    3.times do |att|
      begin
        driver = build_driver
        
        log "Running dream_home_register for #{user} - Attempt #{att+1}"
        dream_home_register(driver, address, birth_day, birth_month, birth_year, cable_provider,
                            city, email, first_name, gender, last_name, phone, state, zip)
        driver.quit
        break
      rescue => ex
        driver.quit
        ex = ex
        next
      end
    end

    raise ex if ex
  end

  def build_driver
    capabilities = Selenium::WebDriver::Remote::Capabilities.new
    capabilities['name'] = 'HGTV Dream Home Register'

    capabilities['platform'] = 'Windows 8.1'
    capabilities['browserName'] = 'Chrome'
    capabilities['version'] = '38'

    capabilities['chromeOptions'] = {'args' => ['--test-type']}

    url = 'http://kangaroo128:19c670aa-47e9-4607-a834-8509f96b186d@ondemand.saucelabs.com:80/wd/hub'
    Selenium::WebDriver.for(:remote,
                            :url => url,
                            :desired_capabilities => capabilities)
  end

  def dream_home_register(driver, address, birth_day, birth_month, birth_year, cable_provider,
                          city, email, first_name, gender, last_name, phone, state, zip)
    driver.navigate.to 'http://www.hgtv.com/design/hgtv-dream-home/sweepstakes/enter'

    Selenium::WebDriver::Wait.new(:timeout => 5).until do
      driver.find_element(:id => 'iFrameResizer0').displayed?
    end
    driver.switch_to.frame('iFrameResizer0')
    driver.find_element(:css => '#login_email input[name=\'email\']').send_keys email
    driver.find_element(:css => 'button.btn_login').click

    Selenium::WebDriver::Wait.new(:timeout => 30).until do
      driver.find_element(:id => 'submit').displayed?
    end

    # Check for welcome back message before continuing.
    if driver.find_elements(:id => 'optin_form').count == 0
      # Fill out 'Your Information' form.
      driver.find_element(:css => '#first_name input[name=\'first_name\']').send_keys first_name
      driver.find_element(:css => '#last_name input[name=\'last_name\']').send_keys last_name
      driver.find_element(:css => '#confirm_email input[name=\'confirm_email\']').send_keys email

      # Address
      driver.find_element(:css => '#address1 input[name=\'address1\']').send_keys address
      driver.find_element(:css => '#city input[name=\'city\']').send_keys city
      select_option(driver, '#state select[name=\'state\']', state)

      driver.find_element(:css => '#zip input[name=\'zip\']').send_keys zip
      driver.find_element(:css => '#phone_number input[name=\'phone_number\']').send_keys phone
      driver.find_element(:css => "#gender input[value='#{gender}']").click

      # Birth Date
      select_option(driver, '#age select[name=\'age.birth_month\']', birth_month)
      select_option(driver, '#age select[name=\'age.birth_day\']', birth_day)
      select_option(driver, '#age select[name=\'age.birth_year\']', birth_year)

      # Cable Provider
      select_option(driver, '#cable_provider select[name=\'cable_provider\']', cable_provider)

      driver.find_element(:css => '#submit button.btn_register').click
    else
      driver.find_element(:css => '#submit button.btn_optin').click
    end

    # Switch back to default frame before searching for success page.
    driver.switch_to.default_content

    # Thank you for entering.
    Selenium::WebDriver::Wait.new(:timeout => 30).until do
      driver.find_element(:css => 'h1.reg_thanks')
    end

    # Enter again on FrontDoor.com------------------------------------------------------------

    driver.navigate.to 'http://www.frontdoor.com/hgtv-dream-home-2015-giveaway?sis=enteragain'

    Selenium::WebDriver::Wait.new(:timeout => 5).until do
      driver.find_element(:id => 'hwframe').displayed?
    end
    driver.switch_to.frame('hwframe')

    driver.find_element(:css => '#login_email input[name=\'email\']').send_keys email
    driver.find_element(:css => '#login_form button.btn_login').click

    driver.switch_to.default_content
    Selenium::WebDriver::Wait.new(:timeout => 5).until do
      driver.find_element(:id => 'hwframe').displayed?
    end
    driver.switch_to.frame('hwframe')

    Selenium::WebDriver::Wait.new(:timeout => 30).until do
      driver.find_element(:id => 'submit').displayed?
    end

    # Check for welcome back message before continuing.
    if driver.find_elements(:id => 'optin_form').count == 0
      # Fill out 'Your Information' form.
      driver.find_element(:css => '#first_name input[name=\'first_name\']').send_keys first_name
      driver.find_element(:css => '#last_name input[name=\'last_name\']').send_keys last_name
      driver.find_element(:css => '#confirm_email input[name=\'confirm_email\']').send_keys email

      # Address
      driver.find_element(:css => '#address1 input[name=\'address1\']').send_keys address
      driver.find_element(:css => '#city input[name=\'city\']').send_keys city
      select_option(driver, '#state select[name=\'state\']', state)

      driver.find_element(:css => '#zip input[name=\'zip\']').send_keys zip
      driver.find_element(:css => '#phone_number input[name=\'phone_number\']').send_keys phone
      driver.find_element(:css => "#gender input[value='#{gender}']").click

      # Birth Date
      select_option(driver, '#age select[name=\'age.birth_month\']', birth_month)
      select_option(driver, '#age select[name=\'age.birth_day\']', birth_day)
      select_option(driver, '#age select[name=\'age.birth_year\']', birth_year)

      # Cable Provider
      select_option(driver, '#cable_provider select[name=\'cable_provider\']', cable_provider)

      driver.find_element(:css => '#submit button.btn_register').click
    else
      driver.find_element(:css => '#submit button.btn_optin').click
    end

    # Switch back to default frame before searching for success page.
    driver.switch_to.default_content

    # Thank you for entering.
    Selenium::WebDriver::Wait.new(:timeout => 30).until do
      driver.find_element(:css => 'div.reg_thanks')
    end
  end

  def select_option(driver, selector, option_text)
    ddl_ele = driver.find_element(:css => selector)
    select = Selenium::WebDriver::Support::Select.new(ddl_ele)
    select.select_by(:text, option_text)
  end

end

WebUiDaemon.mainloop
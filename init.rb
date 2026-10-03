# frozen_string_literal: true

# Copyright 2026 kazuihitoshi@gmail.com
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#     http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.

require 'redmine'

module RedmineExcelMarkdownPaste
  class ViewHook < Redmine::Hook::ViewListener
    def view_layouts_base_html_head(context = {})
      return ''.html_safe unless Setting.text_formatting == 'common_mark'

      controller = context[:controller]
      return ''.html_safe unless controller.respond_to?(:controller_name) && controller.controller_name == 'issues'

      javascript_include_tag('excel_markdown_paste', plugin: 'redmine_excel_markdown_paste')
    end
  end
end

Redmine::Plugin.register :redmine_excel_markdown_paste do
  name 'Excel Markdown Paste'
  author 'kazuihitoshi@gmail.com'
  description 'Pastes an Excel cell range into issue descriptions and notes as a Markdown table when text formatting is Markdown.'
  version '1.0.0'
  requires_redmine version_or_higher: '5.0.0'
end

plugin = Redmine::Plugin.find(:redmine_excel_markdown_paste)

def plugin.name
  I18n.t(:excel_markdown_paste_name, default: 'Excel Markdown Paste')
end

def plugin.description
  I18n.t(:excel_markdown_paste_description, default: 'Pastes an Excel cell range into issue descriptions and notes as a Markdown table when text formatting is Markdown.')
end

# frozen_string_literal: true

# Copyright 2026 kazuihitoshi@gmail.com
#
# This program is free software; you can redistribute it and/or
# modify it under the terms of the GNU General Public License
# as published by the Free Software Foundation; either version 2
# of the License, or (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program; if not, write to the Free Software
# Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301, USA.

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

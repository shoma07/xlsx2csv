# frozen_string_literal: true

require 'bundler/gem_tasks'

require 'rubocop/rake_task'
RuboCop::RakeTask.new(:lint) do |t|
  t.options = %w[--parallel]
end
namespace :lint do
  desc 'Lint fix (Rubocop)'
  task fix: :auto_correct
end

task default: %i[lint]
